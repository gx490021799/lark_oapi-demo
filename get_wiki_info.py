import json
import lark_oapi as lark
from lark_oapi.api.drive.v1 import *

def get_file_statistics(file_token: str, file_type: str, user_access_token: str) -> int:
    # 创建client
    client = lark.Client.builder() \
        .enable_set_token(True) \
        .log_level(lark.LogLevel.INFO) \
        .build()

    # 构造请求对象
    request = GetFileStatisticsRequest.builder() \
        .file_token(file_token) \
        .file_type(file_type) \
        .build()

    # 发起请求
    option = lark.RequestOption.builder().user_access_token(user_access_token).build()
    while True:
        response = client.drive.v1.file_statistics.get(request, option)

        # 处理失败返回
        if not response.success():
            lark.logger.error(
                f"client.drive.v1.file_statistics.get failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}, resp: \n{json.dumps(json.loads(response.raw.content), indent=4, ensure_ascii=False)}")
            return 0  # 返回0表示失败
        else:
            print(f"client.drive.v1.file_statistics.get success, log_id: {response.get_log_id()}")
            # break

    # 保存结果到文件
    file_statistics_json = lark.JSON.marshal(response.data, indent=4)
    with open("file_statistics.json", "w", encoding="utf-8") as f:
        f.write(file_statistics_json)

    # 从文件中读取like_count字段
    with open("file_statistics.json", "r", encoding="utf-8") as f:
        data = json.load(f)
    
    like_count = data.get("statistics", {}).get("like_count", 0)
    return like_count

if __name__ == "__main__":
    file_token = "XWPhwqHkRi8ozakxRb6cz0SknLd"
    file_type = "wiki"
    user_access_token = "u-eVnetEXzl4hE82KA5MprEqhl2VPM10_1po0011m82aPy"
    
    like_count = get_file_statistics(file_token, file_type, user_access_token)
    print(f"like_count: {like_count}")
