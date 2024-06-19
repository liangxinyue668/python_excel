import openpyxl

def PrintTree(level):
    if level == 5:
        return
    print('A' * level + 'B')
    PrintTree(level + 1)
    PrintTree(level + 1)

PrintTree(0)
