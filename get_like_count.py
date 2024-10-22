import json
import pandas as pd
from openpyxl import load_workbook
import lark_oapi as lark
from lark_oapi.api.bitable.v1 import *
from lark_oapi.api.drive.v1 import *


# 多维表格
app_token = "LTDBbgA6Ba5lKDs1pdPcKEpYnoe"
table_id = "tbl8ohS3lZ6ynrbO"
view_id = "vewFAcp4nF"
user_id_type = "user_id"
field_names = ["团队名称", "说明文档", "点赞人数"]

# 云文档
file_token = "XWPhwqHkRi8ozakxRb6cz0SknLd"
file_type = "wiki"
user_access_token = "u-eF5LBteut1097g62WlRbeWhl2rhM10vHN8000l282bDS"
# 定义文档类型变量
valid_document_types = {"sheet", "mindnote", "bitable", "wiki", "file", "docx"}
total_records = 0  # 全局总请求计数

def create_client_and_request(app_token, table_id, user_id_type, view_id, field_names):
    client = lark.Client.builder() \
        .app_id("cli_a7858c7289b0500e") \
        .app_secret("pSLeqUcyr7zlK81mgOTVgd3tg8KR5EZX") \
        .log_level(lark.LogLevel.INFO) \
        .build()

    request = SearchAppTableRecordRequest.builder() \
        .app_token(app_token) \
        .table_id(table_id) \
        .user_id_type(user_id_type) \
        .page_size(500) \
        .request_body(SearchAppTableRecordRequestBody.builder()
            .view_id(view_id)
            .field_names(field_names)  # 包含需要的字段
            .automatic_fields(False)
            .build()) \
        .build()

    return client, request

def main():
    global total_records  # 声明使用全局变量

    client, request = create_client_and_request(app_token, table_id, user_id_type, view_id, field_names)

    # 发起请求
    response: SearchAppTableRecordResponse = client.bitable.v1.app_table_record.search(request)

    # 处理失败返回
    if not response.success():
        print("bitable.v1.app_table_record.search failed, code: %s, msg: %s, log_id: %s")
        lark.logger.error(
            f"client.bitable.v1.app_table_record.search failed, code: {response.code}, msg: {response.msg}, "
            f"log_id: {response.get_log_id()}, resp: \n{json.dumps(json.loads(response.raw.content), indent=4, ensure_ascii=False)}"
        )
        return

    # 保存结果到文件
    result_json = lark.JSON.marshal(response.data, indent=4)
    with open("result.json", "w", encoding="utf-8") as f:
        f.write(result_json)

    # 从文件读取结果并转换为表格
    with open("result.json", "r", encoding="utf-8") as f:
        data = json.load(f)

    # 提取字段并构建数据框
    records = []
    for item in data.get("items", []):
        fields = item.get("fields", {})
        team_name = fields.get("团队名称", [{}])[0].get("text", "")
        

        link = fields.get("说明文档", {}).get("link", "")
        
        # 提取链接中的特定信息
        file_token = link.split("/")[-1].split("?")[0] if link else ""
        
        document_type = link.split("/")[-2] if link else ""

        # 判断link_id和document_type是否为空
        if not file_token or not document_type or document_type not in valid_document_types:
            like_count = "ERR"  # 设置file_info为空
        else:
            like_count = get_file_statistics(file_token, document_type, user_access_token)

        print(f"team_name: {team_name}, file_token: {file_token}, document_type: {document_type}, like_count: {like_count}")
        total_records += 1  # 请求计数增加
        record = {
            "团队名称": team_name,
            "说明文档": link,
            "链接ID": file_token,
            "文档类型": document_type,
            "点赞人数": like_count,
        }
        records.append(record)

    # 创建DataFrame
    df = pd.DataFrame(records)

    # 保存为Excel文件
    # TODO 在打开Excel文件时，也能正常保存数据
    excel_path = "result.xlsx"
    df.to_excel(excel_path, index=False)

    # 设置列宽自适应
    workbook = load_workbook(excel_path)
    sheet = workbook.active

    for column in sheet.columns:
        max_length = 0
        column_cells = [cell for cell in column]
        
        for cell in column_cells:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        
        header_length = len(column_cells[0].value) if column_cells[0].value else 0
        adjusted_width = max(max_length, header_length) + 2
        sheet.column_dimensions[column_cells[0].column_letter].width = adjusted_width

    # 保存调整后的文件
    workbook.save(excel_path)

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
    response = client.drive.v1.file_statistics.get(request, option)

    # 处理失败返回
    if not response.success():
        lark.logger.error(
            f"client.drive.v1.file_statistics.get failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}, resp: \n{json.dumps(json.loads(response.raw.content), indent=4, ensure_ascii=False)}")
        return "ERR"  # 返回0表示失败

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
    main()
    print("total_records: ", total_records)
