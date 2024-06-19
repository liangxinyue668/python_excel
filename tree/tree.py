import openpyxl

id = 0

def PrintTree(level, father):
    global id
    if level == 5:
        return
    title = 'A' * level + 'B'
    id += 1
    me = id
    line = '子任务,%4d,%4d,%s,刘姿彤,周国栋,TR3' % (id, father, title)
    print(line)
    PrintTree(level + 1, me)
    PrintTree(level + 1, me)

print('类型,工作项ID,父级,摘要,经办人,报告人,迭代')
PrintTree(0, 0)
