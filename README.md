# ModifyTestTime

这个工具的功能主要是处理 Excel 文件，特别是针对“测试时间”列的修改，具体包括以下几个方面：
### 主要功能：

1. **读取 Excel 文件**：
   - 工具能够加载当前目录下的 `appList.xlsx` 文件，并读取其中的数据。

2. **查找“测试时间”列**：
   - 工具会自动查找 Excel 表格中名为“测试时间”的列。如果未找到该列，工具会给出明确的提示。

3. **用户输入新时间**：
   - 用户可以输入一个新的测试时间，支持 0 或正整数作为有效输入。如果输入不符合要求，工具会提示用户重新输入。

4. **修改测试时间**：
   - 一旦用户输入有效的时间，工具会将所有“测试时间”列的数据更新为新的输入值。

5. **结果反馈**：
   - 工具会在修改完成后，提示用户修改的成功或失败，并同时输出修改时的日期和时间信息。

6. **程序退出**：
   - 在完成所有操作后，用户可以按任意键退出程序，方便快捷。

### 使用场景：
- 适用于需要批量更新测试时间的场景，例如软件测试、项目管理等，提升工作效率。

这个工具旨在通过简单的交互式界面，使得用户能够快速、方便地更新 Excel 文件中的特定数据。