import json

import lark_oapi as lark
from lark_oapi.api.bitable.v1 import *


# SDK 使用说明: https://github.com/larksuite/oapi-sdk-python#readme
# 以下示例代码是根据 API 调试台参数自动生成，如果存在代码问题，请在 API 调试台填上相关必要参数后再使用
# 复制该 Demo 后, 需要将 "YOUR_APP_ID", "YOUR_APP_SECRET" 替换为自己应用的 APP_ID, APP_SECRET.
def main():
    # 创建client
    client = lark.Client.builder() \
        .app_id("cli_a7858c7289b0500e") \
        .app_secret("pSLeqUcyr7zlK81mgOTVgd3tg8KR5EZX") \
        .log_level(lark.LogLevel.DEBUG) \
        .build()

    # 构造请求对象
    request: SearchAppTableRecordRequest = SearchAppTableRecordRequest.builder() \
        .app_token("LTDBbgA6Ba5lKDs1pdPcKEpYnoe") \
        .table_id("tblaemJDYakC2MTU") \
        .user_id_type("open_id") \
        .page_size(500) \
        .request_body(SearchAppTableRecordRequestBody.builder()
            .view_id("vewn3fQLNZ")
            .field_names(["说明文档"])
            .automatic_fields(False)
            .build()) \
        .build()

    # 发起请求
    while True:
        response: SearchAppTableRecordResponse = client.bitable.v1.app_table_record.search(request)

        # 处理失败返回
        if not response.success():
            lark.logger.error(
                f"client.bitable.v1.app_table_record.search failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}, resp: \n{json.dumps(json.loads(response.raw.content), indent=4, ensure_ascii=False)}")
            return
        else:
            print(f"client.drive.v1.file_statistics.get success, log_id: {response.get_log_id()}")

    # 处理业务结果
    lark.logger.info(lark.JSON.marshal(response.data, indent=4))


if __name__ == "__main__":
    main()
