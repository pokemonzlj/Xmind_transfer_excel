# Xmind_transfer_excel
将用xmind编写的测试用例转化成excel表格形式  
可以根据自己的用例层级结构适当调整递归和输出样式  
# Xmind_transfer_new  
1.xmind用例格式必须跟范本相同，层级以子模块作为根节点，- 思维导图按 5 层（推荐）或 6 层结构进行编写，根节点为当前的 2 个项目，2 级节点为系统下子模块，3 级节点为功能模块，4 级节点为功能点，5 级节点为用例，内容对应用例的概述，要求：
  - 使用清晰、具有描述性的名称，以便于理解和识别。
  - 包含动词，描述测试的行为或操作。
  - 避免使用含糊不清或模棱两可的词汇。
  6 级节点为用例对应的前置条件、步骤、预期，要求：
  - 一个用例下关联的前置条件、步骤、预期各自只能有一条
  - 内容必须以“前置条件：”、“步骤：”、“预期：”的格式开头
2.直接python运行xmind_transter_new.py，选择对应的xmind文件即可生成xlsx文件，存放位置和py文件同目录  
3.在生成的xlsx文件中自行补充前置条件、步骤、预期，完成后再将内容拷贝至导入模版.csv中进行禅道测试用例导入  
