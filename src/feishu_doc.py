# feishu_doc.py
# -*- coding: utf-8 -*-
import json
import logging
import os
import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Dict, Any, Optional

import lark_oapi as lark
from lark_oapi.api.docx.v1 import *
from lark_oapi.api.drive.v1 import DeleteFileRequest
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
        """检查配置是否完整"""
        return bool(self.app_id and self.app_secret)

    def create_daily_doc(self, title: str, content_md: str) -> Optional[str]:
        """
        创建日报文档

        流程：
        1. 清理过期文档（若配置了 FEISHU_DOC_RETENTION_DAYS）
        2. 创建新文档并写入内容
        3. 记录新文档到注册表
        """
        if not self.client or not self.is_configured():
            logger.warning("飞书 SDK 未初始化或配置缺失，跳过创建")
            return None

        # 0. 先清理过期文档（不影响新文档创建）
        self._cleanup_expired_docs()

        try:
            # 1. 创建文档
            # 使用官方 SDK 的 Builder 模式构造请求
            # folder_token 可选：不传则存到应用"我的空间"根目录
            body_builder = CreateDocumentRequestBody.builder() \
                .title(title)
            if self.folder_token:
                body_builder = body_builder.folder_token(self.folder_token)
            create_request = CreateDocumentRequest.builder() \
                .request_body(body_builder.build()) \
                .build()

            response = self.client.docx.v1.document.create(create_request)

            if not response.success():
                logger.error(f"创建文档失败: {response.code} - {response.msg} - {response.error}")
                return None

            doc_id = response.data.document.document_id
            # 这里的 domain 只是为了生成链接，实际访问会重定向
            doc_url = f"https://feishu.cn/docx/{doc_id}"
            logger.info(f"飞书文档创建成功: {title} (ID: {doc_id})")

            # 1.5 记录新文档到注册表（用于后续自动清理）
            self._record_doc(doc_id, doc_url, title)

            # 2. 解析 Markdown 并写入内容
            # 将 Markdown 转换为 SDK 需要的 Block 对象列表
            blocks = self._markdown_to_sdk_blocks(content_md)

            # 飞书 API 限制每次写入 Block 数量（建议 50 个左右），分批写入
            batch_size = 50
            doc_block_id = doc_id  # 文档本身也是一个 block

            for i in range(0, len(blocks), batch_size):
                batch_blocks = blocks[i:i + batch_size]

                # 构造批量添加块的请求
                batch_add_request = CreateDocumentBlockChildrenRequest.builder() \
                    .document_id(doc_id) \
                    .block_id(doc_block_id) \
                    .request_body(CreateDocumentBlockChildrenRequestBody.builder()
                                  .children(batch_blocks)  # SDK 需要 Block 对象列表
                                  .index(-1)  # 追加到末尾
                                  .build()) \
                    .build()

                write_resp = self.client.docx.v1.document_block_children.create(batch_add_request)

                if not write_resp.success():
                    logger.error(f"写入文档内容失败(批次{i}): {write_resp.code} - {write_resp.msg}")

            logger.info(f"文档内容写入完成")

            # 3. 设置文档权限为组织内可读（群成员可打开）
            self._set_doc_public(doc_id)

            return doc_url

        except Exception as e:
            logger.error(f"飞书文档操作异常: {e}")
            import traceback
            logger.error(traceback.format_exc())
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
        """将文档设置为组织内任何人可通过链接查看。"""
        if self.client is None:
            return False
        try:
            from lark_oapi.api.drive.v1 import (
                PatchPermissionPublicRequest,
                PermissionPublicRequest,
            )

            req = (
                PatchPermissionPublicRequest.builder()
                .token(doc_id)
                .type("docx")
                .request_body(
                    PermissionPublicRequest.builder()
                    .link_share_entity("tenant_readable")
                    .build()
                )
                .build()
            )
            resp = self.client.drive.v1.permission_public.patch(req)
            if resp.success():
                logger.info("文档权限已设为组织内可读: %s", doc_id)
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
        表格转换为键值对列表格式。
        """
        from lark_oapi.api.docx.v1 import TextStyle

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
        """将 Markdown 表格行转换为键值对列表并追加到 *blocks*。"""
        from lark_oapi.api.docx.v1 import TextStyle

        if not table_buffer:
            return

        # 解析表格
        rows: List[List[str]] = []
        for line in table_buffer:
            cells = [c.strip() for c in line.split("|")]
            cells = [c for c in cells if c]  # 去掉首尾空串
            if not cells:
                continue
            # 跳过分隔行（如 |---|---|）
            if all(re.match(r"^:?-{2,}:?$", c) for c in cells):
                continue
            rows.append(cells)

        if len(rows) < 2:
            # 单行表格，当成普通文本
            for line in table_buffer:
                elements = self._parse_inline_to_elements(line)
                text_obj = Text.builder().elements(elements).style(TextStyle.builder().build()).build()
                blocks.append(Block.builder().block_type(2).text(text_obj).build())
            return

        header = rows[0]
        data_rows = rows[1:]

        # 先加一个空行作为视觉分隔
        empty = Text.builder().elements(
            self._parse_inline_to_elements("")
        ).style(TextStyle.builder().build()).build()
        blocks.append(Block.builder().block_type(2).text(empty).build())

        for row in data_rows:
            # 补齐列数
            while len(row) < len(header):
                row.append("")
            parts = []
            for idx, cell in enumerate(row):
                key = header[idx] if idx < len(header) else f"列{idx + 1}"
                # 将 cell 内容中的 markdown 转为飞书行内格式
                cell_elements = self._parse_inline_to_elements(cell)
                # 将 key 加粗
                key_style = TextElementStyle.builder().bold(True).build()
                key_tr = TextRun.builder().content(key).text_element_style(key_style).build()
                key_el = TextElement.builder().text_run(key_tr).build()

                colon_tr = TextRun.builder().content("：").text_element_style(
                    TextElementStyle.builder().build()
                ).build()
                colon_el = TextElement.builder().text_run(colon_tr).build()

                all_elements = [key_el, colon_el] + cell_elements
                text_obj = Text.builder().elements(all_elements).style(
                    TextStyle.builder().build()
                ).build()
                blocks.append(Block.builder().block_type(2).text(text_obj).build())

        # 空行结尾
        blocks.append(Block.builder().block_type(2).text(empty).build())