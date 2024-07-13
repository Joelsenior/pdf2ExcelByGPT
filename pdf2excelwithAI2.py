#pdf裁剪
"""
识别章节,然后按照章节对pdf裁剪处理,裁剪后实现pdf的保存。并返回pdf名称。
"""
import fitz
import os
from collections import deque
def split_pdf_by_toc(input_pdf_path, output_dir):
    # Open the input PDF
    pdf_document = fitz.open(input_pdf_path)
    
    # Get the table of contents (TOC)
    toc = pdf_document.get_toc()
    
    if not toc:
        print("No TOC found in the PDF.")
        return
    
    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # # List to store the names of the generated PDFs
    # pdf_filenames = []
    # 使用 deque 来存储生成的 PDF 文件名，初始化为空
    filenames_deque = deque()

    # Process the TOC entries
    for i, entry in enumerate(toc):
        # TOC entry format: [level, title, page number]
        level, title, start_page = entry
        
        # Determine the end page of the chapter
        if i + 1 < len(toc):
            end_page = toc[i + 1][2] - 2
        else:
            end_page = pdf_document.page_count - 1
        
        # Create a new PDF for the chapter
        chapter_pdf = fitz.open()
        chapter_pdf.insert_pdf(pdf_document, from_page=start_page - 1, to_page=end_page)
        
        # Sanitize the chapter title to create a valid filename
        chapter_filename = f"{i + 1}_{title}.pdf"
        chapter_filename = "".join([c if c.isalnum() or c in "._- " else "_" for c in chapter_filename])
        chapter_path = os.path.join(output_dir, chapter_filename)
        
        # Save the chapter PDF
        chapter_pdf.save(chapter_path)
        chapter_pdf.close()
        
        # Add the filename to the list
        filenames_deque.append(chapter_filename)
        
        print(f"Saved chapter: {chapter_filename}")
    
    #写入本地
    with open("filename_deque.txt", "w", encoding='utf-8') as file:
        for item in filenames_deque:
            file.write("%s\n" % item)

    pdf_document.close()
    
    return filenames_deque

# # Example usage
# input_pdf = r"marker-api-master/2.pdf"
# output_directory = "marker-api-master/result"
# split_pdf_by_toc(input_pdf, output_directory)


#pdf识别
"""
利用markerAPI实现pdf的识别,返回json。
需要循环输出
"""
import requests
import os
import json
def recognizePDFByMarker(filename):
    # url = "https://35da-34-83-120-179.ngrok-free.app/convert"
    url = "http://127.0.0.1:8000/convert"
    # for filename in file_path_list:
    #win
    # pdf_file_path = os.path.join(r"C:\Users\lenovo\Desktop\2024\readPDF\marker-api-master\result\",filename)

    # 拼接路径，使用 os.path.join
    pdf_file_path = os.path.join(r"marker-api-master/result", filename)
    # 读取PDF文件内容
    with open(pdf_file_path, 'rb') as pdf_file:
        pdf_content = pdf_file.read()

    # 准备上传的文件数据
    files = {'pdf_file': (os.path.basename(pdf_file_path), pdf_content, 'application/pdf')}

    # 发送POST请求
    response = requests.post(url, files=files)

    # 将JSON响应保存到文件
    json_response = response.json()
    json_file_path = "response.json"
    
    with open(json_file_path, 'w') as json_file:
        json.dump(json_response, json_file)
    
    
    
    print(f"JSON response saved to {json_file_path}")
    
    #提取json中markdown字段
    extractMarkdown = json_response[0]["markdown"]
    # print(extractMarkdown)
    # 将Markdown内容写入到.md文件中
    with open('output_document.md', 'w', encoding='utf-8') as md_file:
        md_file.write(extractMarkdown)

    print("Markdown内容已成功写入到output_document.md文件中。")
    
    return extractMarkdown
        # extractMarkdownList.append(extractMarkdown)
    

# ###实例
# file_path_list = ['6_第一章 总论.pdf']
# recognizePDF(file_path_list=file_path_list)


#OPENAI读取Markdown

from openai import AzureOpenAI
"""
利用大模型接口,错误修正、输出markdown形式excel
"""
def GPTforMarkerResult(markerResult):
    #azure
    API_KEY = "yourkey" 
    # filename_list = recog_result_and_its_filename[1]
    # print(f"markdown识别结果:{markerResult}")
    client = AzureOpenAI(
        api_key=API_KEY,  
        api_version="2024-02-01",
        azure_endpoint = "https://talkaiinstance.openai.azure.com/"
        )
    completion = client.chat.completions.create(
        model="talkAI4O",
        messages=[
            {"role": "system", "content": "你是一个阅读markdown助手，并且帮我从markdown总结内容。具体的：请读取该markdown，理解后将选择题题目及对应的答案找出来,并纠正一些错误的markdown。要求把选择题目、选项和答案填入表格。要求如下：1.表格行为题目序号 2.表格列1为题目标题,列2-列n为选项A到选项n的内容,列n+1为答案列，请一步一步推理,内容输出为可以转化为Excel的markdown。请仔细检查，不包含非表格内容的字符串，仅输出表格的markdown"},
            {"role": "user", "content": markerResult},
        ],
        temperature= 0.2
    )
    result = completion.choices[0].message.content
    print(f"提取内容：{result}")
    # return result,filename_list
    return result
    
# #实例
# file_path_list = ['6_第一章 总论.pdf']
# ans =recognizePDF(file_path_list=file_path_list)
# GPTForJSON(ans)
    


# #markdown转存excel
# """
# 将markdown转化为excel
# """
import pandas as pd
import re

def markdown_to_csv(table_of_markdown,filename):
    # 使用正则表达式匹配Markdown表格中的行和列
    # markdown_table = markdown_table_and_filename
    # filenamelist = markdown_table_and_filename[1]
    # print(type(filenamelist))
    try:
        lines = table_of_markdown.strip().split('\n')
        headers = [h.strip() for h in re.split(r'\s*\|\s*', lines[0].strip('|')) if h.strip()]
        data = []
        
        for line in lines[1:]:
            if line.strip() and not line.strip().startswith('|---'):  # 排除分隔行
                row = [cell.strip() for cell in re.split(r'\s*\|\s*', line.strip('|')) if cell.strip()]
                data.append(row)
        
        # 创建DataFrame
        df = pd.DataFrame(data, columns=headers)
        
        # 清理列名和数据中的空格
        df.columns = df.columns.str.strip()
        df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)  # 使用 apply 替换 applymap
    
        # 将DataFrame存储为CSV文件
        # filename  = filenamedeque.popleft()
        df.to_csv(filename, index=False)
        print("写入成功！")

    except Exception as e:
        print(f"An error occurred in csv: {e}")
        # filenamedeque.popleft()

 


###主函數
#实例
# file_path_list = ['2_版权.pdf']
# file_path_list = deque(['2_版权.pdf'])
if __name__ == "__main__":

    # input_pdf = r"/media/xdu/16445866-4d6f-4702-9066-ca085e6b07dd/2T/readPDF/marker-api-master/2.pdf"
    # output_directory = "marker-api-master/result"

    # file_name_deque =split_pdf_by_toc(input_pdf, output_directory)
    
    # 从文本文件读取并创建新的deque
    # with open("filename_deque.txt", "r", encoding='utf-8') as file:
    #     lines = file.readlines()

    # # 去除每行末尾的换行符并创建deque
    # file_name_deque = deque(line.rstrip() for line in lines)

    # # 打印读取的deque内容
    # print(list(file_name_deque))

    file_name_deque = (["126_第一章 总论.pdf"])
    for filename in file_name_deque:
        #marker识别结果
        ans =recognizePDFByMarker(filename)
        #gpt识别结果
        markdown_table = GPTforMarkerResult(ans)
        # # gpt结果转化为csv
        markdown_to_csv(markdown_table,filename)

    print("完成写入！")
