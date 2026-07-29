# feishu_doc.py
# -*- coding: utf-8 -*-
import io
import json
import logging
import os
import re
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Dict, Any, Optional

import lark_oapi as lark
from lark_oapi.api.docx.v1 import *
from lark_oapi.api.drive.v1 import (
    CreateFolderFileRequest,
    CreateFolderFileRequestBody,
    CreateImportTaskRequest,
    DeleteFileRequest,
    GetImportTaskRequest,
    ImportTask,
    ImportTaskMountPoint,
    UploadAllFileRequest,
    UploadAllFileRequestBody,
)
from src.config import get_config

logger = logging.getLogger(__name__)

# 文档注册表路径（记录已创建的文档 ID，用于定期清理）
_DEFAULT_REGISTRY_PATH = Path(__file__).parent.parent / "data" / "feishu_doc_registry.json"


class FeishuDocManager:
    """飞书云文档管理器 (基于官方 SDK lark-oapi)"""

    def __init__(self):
        self.config = get_config()
        self.app_id = self.config.feishu_app_id
        self.app_secret = self.config.feishu_app_secret
        self.folder_token = self.config.feishu_folder_token
        self._retention_days = int(getattr(self.config, "feishu_doc_retention_days", 0) or 0)
        self._registry_path = _DEFAULT_REGISTRY_PATH
        self._pending_tables: List = []  # (header, rows) accumulated during markdown parsing

        # 解析域名：feishu(飞书国内) / lark(国际版)
        raw_domain = (
            getattr(self.config, "feishu_domain", None)
            or __import__("os").getenv("FEISHU_DOMAIN", "feishu")
        ).strip().lower()
        if raw_domain not in ("feishu", "lark"):
            raw_domain = "feishu"
        try:
            from lark_oapi.core.const import FEISHU_DOMAIN as _SDK_FEISHU
            from lark_oapi.core.const import LARK_DOMAIN as _SDK_LARK
            self._feishu_domain = _SDK_FEISHU if raw_domain == "feishu" else _SDK_LARK
        except ImportError:
            self._feishu_domain = None

        # 初始化 SDK 客户端
        # SDK 会自动处理 tenant_access_token 的获取和刷新，无需人工干预
        if self.is_configured():
            builder = lark.Client.builder() \
                .app_id(self.app_id) \
                .app_secret(self.app_secret) \
                .log_level(lark.LogLevel.INFO)
            if self._feishu_domain is not None:
                builder = builder.domain(self._feishu_domain)
            self.client = builder.build()
        else:
            self.client = None

    def is_configured(self) -> bool:
        """检查配置是否完整（需要 FEISHU_APP_ID + FEISHU_APP_SECRET，文件夹自动创建）"""
        return bool(self.app_id and self.app_secret)

    def _ensure_folder_token(self) -> Optional[str]:
        """确保有应用可写入的文件夹 token。

        优先级：
        1. 使用 .env 中配置的 FEISHU_FOLDER_TOKEN
        2. 使用注册表中缓存的自动创建的文件夹 token
        3. 通过 API 自动创建一个应用所属的文件夹

        Returns:
            有效的文件夹 token，失败返回 None。
        """
        # 1. 检查配置的 folder_token
        if self.folder_token:
            return self.folder_token

        # 2. 检查注册表中的缓存
        registry = self._load_registry()
        cached_token = registry.get("auto_folder_token", "")
        if cached_token:
            logger.debug("使用注册表中缓存的文件夹 token: %s", cached_token)
            self.folder_token = cached_token
            return cached_token

        # 3. 通过 API 在应用根目录自动创建文件夹（folder_token='' 表示根目录）
        try:
            body = (
                CreateFolderFileRequestBody.builder()
                .name("daily_reports")
                .folder_token("")
                .build()
            )
            req = (
                CreateFolderFileRequest.builder()
                .request_body(body)
                .build()
            )
            resp = self.client.drive.v1.file.create_folder(req)
            if resp.success() and resp.data and resp.data.token:
                new_token = resp.data.token
                registry["auto_folder_token"] = new_token
                self._save_registry(registry)
                self.folder_token = new_token
                logger.info("自动创建日报文件夹成功: token=%s", new_token)
                return new_token
            logger.error(
                "自动创建文件夹失败: code=%s msg=%s",
                getattr(resp, "code", "?"), getattr(resp, "msg", "?"),
            )
        except Exception as e:
            logger.error("自动创建文件夹异常: %s", e)

        return None

    def create_daily_doc(self, title: str, content_md: str) -> Optional[str]:
        """
        创建日报文档（Markdown 导入方式，效果等同于在飞书文档中粘贴 Markdown）。

        流程：
        1. 清理过期文档（若配置了 FEISHU_DOC_RETENTION_DAYS）
        2. 上传 Markdown 文件到飞书云盘
        3. 通过 import_task API 将 Markdown 导入为飞书文档
        4. 轮询等待导入完成
        5. 记录新文档到注册表
        6. 设置文档权限为公开可读
        """
        if not self.client or not self.is_configured():
            logger.warning("飞书 SDK 未初始化或配置缺失，跳过创建")
            return None

        # 0. 确保文件夹 token 有效（没有则自动创建一个应用所属的文件夹）
        folder_token = self._ensure_folder_token()
        if not folder_token:
            logger.error("无法获取有效的文件夹 token，跳过创建")
            return None

        # 1. 先清理过期文档（不影响新文档创建）
        self._cleanup_expired_docs()

        try:
            # 2. 上传 Markdown 内容为文件
            content_bytes = content_md.encode("utf-8")
            file_obj = io.BytesIO(content_bytes)
            upload_body = (
                UploadAllFileRequestBody.builder()
                .file_name(f"{title}.md")
                .parent_type("explorer")
                .size(len(content_bytes))
                .file(file_obj)
                .build()
            )
            upload_resp = self.client.drive.v1.file.upload_all(
                UploadAllFileRequest.builder().request_body(upload_body).build()
            )
            if not upload_resp.success():
                logger.error(
                    "上传 Markdown 文件失败: code=%s msg=%s",
                    upload_resp.code, upload_resp.msg,
                )
                return None
            file_token = upload_resp.data.file_token
            logger.debug("Markdown 文件上传成功, file_token=%s", file_token)

            # 3. 创建导入任务（md → docx）
            # 注意：飞书 API 要求 mount_type=1 + mount_key（文件夹 token）
            mp = (
                ImportTaskMountPoint.builder()
                .mount_type(1)
                .mount_key(folder_token)
                .build()
            )
            import_task = (
                ImportTask.builder()
                .file_token(file_token)
                .file_extension("md")
                .type("docx")
                .file_name(title)
                .point(mp)
                .build()
            )
            import_resp = self.client.drive.v1.import_task.create(
                CreateImportTaskRequest.builder().request_body(import_task).build()
            )
            if not import_resp.success():
                err_detail = ""
                if import_resp.error and hasattr(import_resp.error, "field_violations"):
                    err_detail = f" violations={import_resp.error.field_violations}"
                logger.error(
                    "创建导入任务失败: code=%s msg=%s%s",
                    import_resp.code, import_resp.msg, err_detail,
                )
                return None
            ticket = import_resp.data.ticket
            logger.debug("导入任务已创建, ticket=%s", ticket)

            # 4. 轮询等待导入完成
            doc_token = self._poll_import_task(ticket)
            if not doc_token:
                logger.error("导入任务未完成，未能获取文档 token")
                return None

            doc_url = f"https://feishu.cn/docx/{doc_token}"
            logger.info("飞书文档创建成功: %s (ID: %s)", title, doc_token)

            # 5. 记录新文档到注册表（用于后续自动清理）
            self._record_doc(doc_token, doc_url, title)

            # 6. 设置文档权限为公开可读（群成员+外部人员均可访问）
            self._set_doc_public(doc_token)

            return doc_url

        except Exception as e:
            logger.error("飞书文档操作异常: %s", e)
            import traceback
            logger.error(traceback.format_exc())
            return None

    def _poll_import_task(self, ticket: str, max_retries: int = 15, interval: float = 1.0) -> Optional[str]:
        """轮询导入任务结果，返回导入完成后的文档 token。

        Args:
            ticket: 导入任务 ticket。
            max_retries: 最大轮询次数（默认 15 次）。
            interval: 每次轮询间隔秒数（默认 1.0 秒）。

        Returns:
            导入成功时返回文档 token，失败/超时返回 None。
        """
        for i in range(max_retries):
            time.sleep(interval)
            get_resp = self.client.drive.v1.import_task.get(
                GetImportTaskRequest.builder().ticket(ticket).build()
            )
            if not get_resp.success():
                logger.warning(
                    "查询导入任务失败 [%d/%d]: code=%s msg=%s",
                    i + 1, max_retries, get_resp.code, get_resp.msg,
                )
                continue

            result = get_resp.data.result
            if result is None:
                logger.warning("查询导入任务返回空结果 [%d/%d]", i + 1, max_retries)
                continue

            job_status = getattr(result, "job_status", None)
            err_msg = getattr(result, "job_error_msg", "") or ""
            logger.debug(
                "导入任务状态 [%d/%d]: job_status=%s, err_msg=%s",
                i + 1, max_retries, job_status, err_msg,
            )

            if job_status == 0:  # 完成
                doc_token = getattr(result, "token", "") or ""
                if doc_token and err_msg == "success":
                    logger.info("导入任务完成, token=%s", doc_token)
                    return doc_token
                logger.warning(
                    "导入任务结束但状态异常: job_status=0, err_msg=%s, token=%s",
                    err_msg, doc_token,
                )
                return None
            elif job_status is not None and job_status < 0:  # 负值表示失败
                logger.error("导入任务失败: job_status=%s, err_msg=%s", job_status, err_msg)
                return None
            # job_status > 0 表示排队/处理中，继续轮询

        logger.error("导入任务轮询超时（%d 次）", max_retries)
        return None

    # ------------------------------------------------------------------
    # Doc registry & auto-cleanup
    # ------------------------------------------------------------------

    def _load_registry(self) -> dict:
        """加载文档注册表。"""
        try:
            if self._registry_path.is_file():
                with open(self._registry_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict) and "docs" in data:
                        return data
        except Exception as e:
            logger.warning("加载飞书文档注册表失败: %s", e)
        return {"docs": []}

    def _save_registry(self, registry: dict) -> None:
        """保存文档注册表。"""
        try:
            self._registry_path.parent.mkdir(parents=True, exist_ok=True)
            tmp = self._registry_path.with_suffix(".tmp")
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(registry, f, ensure_ascii=False, indent=2)
            os.replace(tmp, self._registry_path)  # atomic on POSIX
        except Exception as e:
            logger.warning("保存飞书文档注册表失败: %s", e)

    def _record_doc(self, doc_id: str, url: str, title: str) -> None:
        """将新创建的文档记录到注册表。"""
        registry = self._load_registry()
        registry.setdefault("docs", []).append({
            "doc_id": doc_id,
            "url": url,
            "title": title,
            "created_at": datetime.now().isoformat(),
        })
        self._save_registry(registry)
        logger.debug("文档已记录: %s (%s)", title, doc_id)

    def _set_doc_public(self, doc_id: str) -> bool:
        """将文档设置为任何人可通过链接查看（组织内外均可）。"""
        if self.client is None:
            return False
        try:
            from lark_oapi.api.drive.v1 import (
                PatchPermissionPublicRequest,
                PermissionPublicRequest,
            )

            body = (
                PermissionPublicRequest.builder()
                .external_access(True)
                .invite_external(True)
                .link_share_entity("anyone_readable")
                .build()
            )
            req = (
                PatchPermissionPublicRequest.builder()
                .token(doc_id)
                .type("docx")
                .request_body(body)
                .build()
            )
            resp = self.client.drive.v1.permission_public.patch(req)
            if resp.success():
                logger.info("文档权限已设为公开可读: %s", doc_id)
                return True
            logger.warning(
                "设置文档权限失败 (%s): code=%s msg=%s",
                doc_id, resp.code, resp.msg,
            )
            return False
        except Exception as e:
            logger.warning("设置文档权限异常 (%s): %s", doc_id, e)
            return False

    def _delete_doc(self, doc_id: str) -> bool:
        """通过 Drive API 删除单个文档。"""
        if self.client is None:
            return False
        try:
            req = DeleteFileRequest.builder().file_token(doc_id).build()
            resp = self.client.drive.v1.file.delete(req)
            if resp.success():
                logger.info("已删除过期飞书文档: %s", doc_id)
                return True
            logger.warning(
                "删除飞书文档失败 (%s): code=%s msg=%s",
                doc_id, resp.code, resp.msg,
            )
            return False
        except Exception as e:
            logger.warning("删除飞书文档异常 (%s): %s", doc_id, e)
            return False

    def _cleanup_expired_docs(self) -> int:
        """删除超过保留天数的文档。

        仅在 ``self._retention_days > 0`` 时执行。
        删除失败的文档会保留在注册表中，下次重试。

        Returns:
            本次删除的文档数量。
        """
        if self._retention_days <= 0:
            return 0

        registry = self._load_registry()
        docs = registry.get("docs", [])
        if not docs:
            return 0

        cutoff = datetime.now() - timedelta(days=self._retention_days)
        remaining = []
        deleted = 0

        for doc in docs:
            doc_id = doc.get("doc_id", "")
            try:
                created_at = datetime.fromisoformat(doc["created_at"])
            except (ValueError, KeyError):
                remaining.append(doc)
                continue

            if created_at < cutoff:
                if self._delete_doc(doc_id):
                    deleted += 1
                    continue
                # 删除失败 -- 保留在注册表中，下次重试
            remaining.append(doc)

        if deleted > 0:
            registry["docs"] = remaining
            self._save_registry(registry)
            logger.info("飞书文档自动清理完成：删除 %d 个过期文档，剩余 %d 个", deleted, len(remaining))

        return deleted

    # ------------------------------------------------------------------
    # Markdown -> Feishu blocks
    # ------------------------------------------------------------------

    @staticmethod
    def _parse_inline_to_elements(text: str) -> List:
        """Parse inline Markdown formatting into Feishu TextElement objects.

        Handles: **bold**, *italic*, `code`, [text](url), plain text.
        """
        from lark_oapi.api.docx.v1 import (
            Link,
            TextElement,
            TextElementStyle,
            TextRun,
        )

        elements: List = []
        i = 0
        n = len(text)

        while i < n:
            # **bold**
            if text[i : i + 2] == "**" and i + 2 < n:
                end = text.find("**", i + 2)
                if end > i:
                    style = TextElementStyle.builder().bold(True).build()
                    tr = TextRun.builder().content(text[i + 2 : end]).text_element_style(style).build()
                    elements.append(TextElement.builder().text_run(tr).build())
                    i = end + 2
                    continue

            # *italic* (single * -- must not be **)
            if text[i] == "*" and (i + 1 >= n or text[i + 1] != "*"):
                end = text.find("*", i + 1)
                if end > i:
                    style = TextElementStyle.builder().italic(True).build()
                    tr = TextRun.builder().content(text[i + 1 : end]).text_element_style(style).build()
                    elements.append(TextElement.builder().text_run(tr).build())
                    i = end + 1
                    continue

            # `inline code`
            if text[i] == "`":
                end = text.find("`", i + 1)
                if end > i:
                    style = TextElementStyle.builder().inline_code(True).build()
                    tr = TextRun.builder().content(text[i + 1 : end]).text_element_style(style).build()
                    elements.append(TextElement.builder().text_run(tr).build())
                    i = end + 1
                    continue

            # [link text](url)
            if text[i] == "[":
                end_bracket = text.find("]", i + 1)
                if end_bracket > i and end_bracket + 1 < n and text[end_bracket + 1] == "(":
                    end_paren = text.find(")", end_bracket + 2)
                    if end_paren > end_bracket:
                        link_text = text[i + 1 : end_bracket]
                        link_url = text[end_bracket + 2 : end_paren]
                        link_obj = Link.builder().url(link_url).build()
                        style = TextElementStyle.builder().link(link_obj).build()
                        tr = TextRun.builder().content(link_text).text_element_style(style).build()
                        elements.append(TextElement.builder().text_run(tr).build())
                        i = end_paren + 1
                        continue

            # Plain text -- scan to next special character
            next_pos = n
            for ch in ("**", "*", "`", "["):
                p = text.find(ch, i + 1 if ch != "**" else i)
                if p != -1 and p < next_pos:
                    next_pos = p
            style = TextElementStyle.builder().build()
            tr = TextRun.builder().content(text[i:next_pos]).text_element_style(style).build()
            elements.append(TextElement.builder().text_run(tr).build())
            i = next_pos

        if not elements:
            tr = TextRun.builder().content("").text_element_style(TextElementStyle.builder().build()).build()
            elements.append(TextElement.builder().text_run(tr).build())
        return elements

    def _markdown_to_sdk_blocks(self, md_text: str) -> List[Block]:
        """将 Markdown 转换为飞书 SDK 的 Block 对象。

        支持：标题（H1-H3）、分割线、无序列表、行内格式（粗体/斜体/行内代码/链接）。
        表格暂存到 _pending_tables，由 _create_real_table 生成真表格。
        """
        from lark_oapi.api.docx.v1 import TextStyle

        self._pending_tables = []
        blocks: List = []
        table_buffer: List[str] = []
        lines = md_text.split("\n")

        for line in lines:
            stripped = line.strip()

            # 缓冲表格行（以 | 开头且以 | 结尾）
            if stripped.startswith("|") and stripped.endswith("|"):
                table_buffer.append(stripped)
                continue

            # 刷新表格缓冲区
            if table_buffer:
                self._flush_table_buffer(table_buffer, blocks)
                table_buffer = []

            if not stripped:
                continue

            # 标题
            if stripped.startswith("# ") and not stripped.startswith("## "):
                text_content = stripped[2:]
                elements = self._parse_inline_to_elements(text_content)
                text_obj = Text.builder().elements(elements).style(TextStyle.builder().build()).build()
                blocks.append(Block.builder().block_type(3).heading1(text_obj).build())
                continue
            if stripped.startswith("## ") and not stripped.startswith("### "):
                text_content = stripped[3:]
                elements = self._parse_inline_to_elements(text_content)
                text_obj = Text.builder().elements(elements).style(TextStyle.builder().build()).build()
                blocks.append(Block.builder().block_type(4).heading2(text_obj).build())
                continue
            if stripped.startswith("### "):
                text_content = stripped[4:]
                elements = self._parse_inline_to_elements(text_content)
                text_obj = Text.builder().elements(elements).style(TextStyle.builder().build()).build()
                blocks.append(Block.builder().block_type(5).heading3(text_obj).build())
                continue

            # 分割线
            if stripped in ("---", "***"):
                div = Divider.builder().build()
                blocks.append(Block.builder().block_type(22).divider(div).build())
                continue

            # 普通文本
            elements = self._parse_inline_to_elements(stripped)
            text_obj = Text.builder().elements(elements).style(TextStyle.builder().build()).build()
            blocks.append(Block.builder().block_type(2).text(text_obj).build())

        # 刷新末尾的表格缓冲区
        if table_buffer:
            self._flush_table_buffer(table_buffer, blocks)

        return blocks

    def _flush_table_buffer(self, table_buffer: List[str], blocks: List) -> None:
        """解析 Markdown 表格，暂存到 _pending_tables 用于后续创建真表格。"""
        if not table_buffer:
            return

        rows: List[List[str]] = []
        for line in table_buffer:
            cells = [c.strip() for c in line.split("|")]
            cells = [c for c in cells if c]
            if not cells:
                continue
            if all(re.match(r"^:?-{2,}:?$", c) for c in cells):
                continue
            rows.append(cells)

        if len(rows) < 2:
            from lark_oapi.api.docx.v1 import TextStyle
            for line in table_buffer:
                elements = self._parse_inline_to_elements(line)
                text_obj = Text.builder().elements(elements).style(TextStyle.builder().build()).build()
                blocks.append(Block.builder().block_type(2).text(text_obj).build())
            return

        header = rows[0]
        data_rows = rows[1:]

        # 宽表（>3列）转竖排键值对，避免真表格被挤扁
        if len(header) > 3:
            self._flush_wide_table_as_kv(header, data_rows, blocks)
            return

        # 窄表（≤3列）暂存，后续创建真表格
        self._pending_tables.append((header, data_rows))
        from lark_oapi.api.docx.v1 import TextStyle
        empty = Text.builder().elements(
            self._parse_inline_to_elements("")
        ).style(TextStyle.builder().build()).build()
        blocks.append(Block.builder().block_type(2).text(empty).build())

    def _flush_wide_table_as_kv(
        self, header: List[str], rows: List[List[str]], blocks: List
    ) -> None:
        """将宽表（>3列）转竖排键值对，每个数据行输出一组加粗标题: 值的行。"""
        from lark_oapi.api.docx.v1 import TextStyle

        empty = Text.builder().elements(
            self._parse_inline_to_elements("")
        ).style(TextStyle.builder().build()).build()
        blocks.append(Block.builder().block_type(2).text(empty).build())

        for row_idx, row in enumerate(rows):
            while len(row) < len(header):
                row.append("")
            # 多行之间加分割
            if row_idx > 0:
                div = Divider.builder().build()
                blocks.append(Block.builder().block_type(22).divider(div).build())

            for col_idx, cell in enumerate(row):
                key = header[col_idx] if col_idx < len(header) else ""
                key_style = TextElementStyle.builder().bold(True).build()
                key_tr = TextRun.builder().content(key).text_element_style(key_style).build()
                key_el = TextElement.builder().text_run(key_tr).build()
                colon_tr = TextRun.builder().content("：").text_element_style(
                    TextElementStyle.builder().build()
                ).build()
                colon_el = TextElement.builder().text_run(colon_tr).build()
                cell_elements = self._parse_inline_to_elements(cell)
                all_el = [key_el, colon_el] + cell_elements
                text_obj = Text.builder().elements(all_el).style(TextStyle.builder().build()).build()
                blocks.append(Block.builder().block_type(2).text(text_obj).build())

        blocks.append(Block.builder().block_type(2).text(empty).build())

    def _create_real_table(
        self, doc_id: str, header: List[str], rows: List[List[str]]
    ) -> None:
        """在文档中创建一个真实的飞书表格并填充单元格内容。"""
        if self.client is None:
            return

        col_count = len(header)
        row_count = len(rows)
        if row_count < 1 or col_count < 1:
            logger.warning("跳过空表格: %d行 %d列", row_count, col_count)
            return

        # 补齐每行列数
        for r in rows:
            while len(r) < col_count:
                r.append("")

        try:
            from lark_oapi.api.docx.v1 import (
                CreateDocumentBlockChildrenRequest,
                CreateDocumentBlockChildrenRequestBody,
                TableProperty,
                TextStyle,
            )
            from lark_oapi.api.docx.v1 import Table as FeishuTable

            # 1. 创建 Table block（服务器自动生成占位单元格）
            table_prop = TableProperty.builder().row_size(row_count).column_size(col_count).header_row(True).build()
            table_obj = FeishuTable.builder().property(table_prop).build()
            table_block = Block.builder().block_type(31).table(table_obj).build()

            req = (
                CreateDocumentBlockChildrenRequest.builder()
                .document_id(doc_id)
                .block_id(doc_id)
                .request_body(
                    CreateDocumentBlockChildrenRequestBody.builder()
                    .children([table_block])
                    .index(-1)
                    .build()
                )
                .build()
            )
            resp = self.client.docx.v1.document_block_children.create(req)
            if not resp.success():
                logger.warning("创建表格块失败: code=%s msg=%s", resp.code, resp.msg)
                return
            if not resp.data or not resp.data.children:
                logger.warning("创建表格块返回空数据，跳过")
                return

            cell_ids = list(resp.data.children[0].children or [])
            if len(cell_ids) < col_count * row_count:
                logger.warning("表格单元格数量不足: 期望 %d 实际 %d", col_count * row_count, len(cell_ids))
                return

            # 2. 逐行填充单元格内容
            for row_idx in range(row_count):
                for col_idx in range(col_count):
                    cell_id = cell_ids[row_idx * col_count + col_idx]
                    cell_text = rows[row_idx][col_idx] if col_idx < len(rows[row_idx]) else ""
                    if row_idx == 0:
                        style = TextElementStyle.builder().bold(True).build()
                        tr = TextRun.builder().content(cell_text).text_element_style(style).build()
                        el = TextElement.builder().text_run(tr).build()
                        cell_elements = [el]
                    else:
                        cell_elements = self._parse_inline_to_elements(cell_text)
                    text_obj = Text.builder().elements(cell_elements).style(TextStyle.builder().build()).build()
                    cell_block = Block.builder().block_type(2).text(text_obj).build()

                    cell_req = (
                        CreateDocumentBlockChildrenRequest.builder()
                        .document_id(doc_id)
                        .block_id(cell_id)
                        .request_body(
                            CreateDocumentBlockChildrenRequestBody.builder()
                            .children([cell_block])
                            .index(-1)
                            .build()
                        )
                        .build()
                    )
                    cell_resp = self.client.docx.v1.document_block_children.create(cell_req)
                    if not cell_resp.success():
                        logger.debug(
                            "表格单元格填充失败 [%d,%d]: code=%s msg=%s",
                            row_idx, col_idx, cell_resp.code, cell_resp.msg,
                        )

            logger.info("表格创建完成: %d 行 %d 列", row_count, col_count)

        except Exception as e:
            logger.warning("表格创建异常 (%dx%d): %s，回退跳过此表格", row_count, col_count, e)