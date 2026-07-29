#!/usr/bin/env python3
"""测试飞书 Markdown 导入 API（效果等同直接粘贴到飞书文档）"""
import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
from src.config import get_config
import lark_oapi as lark
from lark_oapi.core.const import FEISHU_DOMAIN as _SDK_FEISHU
from lark_oapi.api.drive.v1 import (
    UploadAllFileRequest,
    UploadAllFileRequestBody,
    CreateImportTaskRequest,
    ImportTask,
    ImportTaskMountPoint,
    PatchPermissionPublicRequest,
    PermissionPublicRequest,
)
from lark_oapi.api.docx.v1 import GetDocumentRequest

cfg = get_config()
client = lark.Client.builder() \
    .app_id(cfg.feishu_app_id).app_secret(cfg.feishu_app_secret) \
    .domain(_SDK_FEISHU).log_level(lark.LogLevel.WARNING).build()

# 读真实报告
report_path = sys.argv[1] if len(sys.argv) > 1 else "reports/report_20260728.md"
with open(report_path) as f:
    content = f.read()

content_bytes = content.encode("utf-8")

# 1. 上传 Markdown 文件
print("1. 上传文件...")
file_obj = io.BytesIO(content_bytes)
upload_body = (
    UploadAllFileRequestBody.builder()
    .file_name("report.md")
    .parent_type("explorer")
    .size(len(content_bytes))
    .file(file_obj)
    .build()
)
upload_resp = client.drive.v1.file.upload_all(
    UploadAllFileRequest.builder().request_body(upload_body).build()
)
if not upload_resp.success():
    print(f"   上传失败: code={upload_resp.code} msg={upload_resp.msg}")
    sys.exit(1)
file_token = upload_resp.data.file_token
print(f"   成功, file_token={file_token}")

# 2. 创建导入任务（md → docx）
# 注意：飞书 API 已变更，mount_type 仅支持 1（导入到指定文件夹），
# mount_key 必须填写有效的文件夹 token。
# 如果没有手动配置，则通过 API 自动创建一个应用专属文件夹。
print("2. 创建导入任务...")
folder_token = getattr(cfg, "feishu_folder_token", None)
if not folder_token:
    from lark_oapi.api.drive.v1 import (
        CreateFolderFileRequest,
        CreateFolderFileRequestBody,
    )
    print("   未配置 FEISHU_FOLDER_TOKEN，自动创建应用专属文件夹...")
    try:
        folder_body = (
            CreateFolderFileRequestBody.builder()
            .name("daily_reports")
            .folder_token("")  # 空串=应用根目录
            .build()
        )
        folder_req = (
            CreateFolderFileRequest.builder()
            .request_body(folder_body)
            .build()
        )
        folder_resp = client.drive.v1.file.create_folder(folder_req)
        if not folder_resp.success():
            print(f"   自动创建文件夹失败: code={folder_resp.code} msg={folder_resp.msg}")
            sys.exit(1)
        folder_token = folder_resp.data.token
        print(f"   自动创建成功, folder_token={folder_token}")
        print(f"   (可将此 token 填入 .env 的 FEISHU_FOLDER_TOKEN 以便复用)")
    except Exception as e:
        print(f"   自动创建文件夹异常: {e}")
        sys.exit(1)
else:
    print(f"   使用配置的 folder_token={folder_token}")

mp = ImportTaskMountPoint.builder().mount_type(1).mount_key(folder_token).build()
import_task = ImportTask.builder() \
    .file_token(file_token) \
    .file_extension("md") \
    .type("docx") \
    .file_name("daily_report_import_test") \
    .point(mp) \
    .build()
import_resp = client.drive.v1.import_task.create(
    CreateImportTaskRequest.builder().request_body(import_task).build()
)
if not import_resp.success():
    err_detail = ""
    if import_resp.error and hasattr(import_resp.error, "field_violations"):
        err_detail = f" violations={import_resp.error.field_violations}"
    print(f"   导入失败: code={import_resp.code} msg={import_resp.msg}{err_detail}")
    sys.exit(1)

# 获取 ticket，轮询导入结果
ticket = import_resp.data.ticket
print(f"   成功, ticket={ticket}")

# 轮询获取导入结果
print("   轮询导入结果...")
from lark_oapi.api.drive.v1 import GetImportTaskRequest
import time
doc_token = ""
for i in range(15):
    time.sleep(1)
    get_resp = client.drive.v1.import_task.get(
        GetImportTaskRequest.builder().ticket(ticket).build()
    )
    if get_resp.success() and get_resp.data.result:
        r = get_resp.data.result
        js = getattr(r, 'job_status', None)
        err_msg = getattr(r, 'job_error_msg', '') or ''
        print(f"   [{i+1}] job_status={js}, err_msg={err_msg}")
        if js == 0:  # 完成
            doc_token = getattr(r, 'token', '') or ''
            if doc_token and err_msg == 'success':
                print(f"   导入完成, token={doc_token}")
                break
            print(f"   导入异常: job_status=0, err_msg={err_msg}")
            sys.exit(1)
        elif js is not None and js < 0:  # 失败
            print(f"   导入失败: {err_msg}")
            print(f"   完整result: {r}")
            sys.exit(1)
        # js > 0 表示排队/处理中，继续轮询
else:
    print("   轮询超时")
    sys.exit(1)

# 3. 设置公开权限
print("3. 设置权限...")
perm_body = (
    PermissionPublicRequest.builder()
    .external_access(True)
    .invite_external(True)
    .link_share_entity("anyone_readable")
    .build()
)
perm_resp = client.drive.v1.permission_public.patch(
    PatchPermissionPublicRequest.builder()
    .token(doc_token)
    .type("docx")
    .request_body(perm_body)
    .build()
)
print(f"   权限: {'OK' if perm_resp.success() else f'FAIL {perm_resp.code}'}")

doc_url = f"https://feishu.cn/docx/{doc_token}"
print(f"\n✅ 完成: {doc_url}")
print("打开链接查看效果（导入方式 = 粘贴 Markdown 效果）")
