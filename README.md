# classArrangeForYijiyizhong
一个用来解决一机一中神奇的走班排课问题的小程序
这是一个需要神奇力量来解决的排课问题，** 1000+名学生，88个班 **,
按照成绩进行分班并且要求
** 不重不漏 **
*********************
目前解决程度：

* 可以为绝大多数学生正确安排课程

* 由于部分学生偏科现象严重，并且A1班只安排一节，在较好的情况下仍有90个学生的物理课不能自动安排
*********************
需要为你的Python3.x准备以下额外模块

* pymongo（使用了一个本地的MongoDB，也就意味着你需要安装MongoDB）

* xlwt和xlrd用来读写成绩单
**********************
文件 成绩数据 和 课程安排数据 是程序可接受的输入格式

为了保护学生个人信息，我修改了成绩数据和人名的对应关系，这可能降低了难度，因为按照姓名音序排列后，偏科现象就不存在了（想想为什么）
**********************
欢迎各位想玩的和我一起把这个程序完善好
