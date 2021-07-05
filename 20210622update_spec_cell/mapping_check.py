import os,re,csv
import xlwings as xw

f_GVL = r'C:\Users\sesa504595\Desktop\MyTool\20210622update_spec_cell\charlie_20210622_1.csv'
f_GVS = r'C:\Users\sesa504595\Desktop\MyTool\20210622update_spec_cell\Jedi_20210622_1.csv'


def get_requirements_dict(file):
    '''
    :param file: caliber dump file
    :return: dict {'68980': ('BMC communication OK', 'WHAT'),...}  先确认encoding
    '''
    Dict = {}
    with open(file) as f:
        reader = csv.reader(f)
        for row in reader:
            # print(row)
            type = re.findall(r'[(](.*?)[)]$',row[2]) #提取结尾小括号中的内容
            if type:
                Dict[row[4]] = (row[3],type[0])
    return Dict

# gvl = get_GVL_requirements_dict(f_GVL)
# with open('Dict_GVL.txt','a') as f:
#     print(gvl,file=f)

def get_map(GVL_D,GVS_D):
    '''
    创建一个一个新的Dict
    根据GVL requirement Number 查找到 Jedi number
    format is {'54618': (('Archived requirements', 'HOW'),'88888','WHAT')...}
    '''
    for key, value in GVL_D.items():
        for k, v in GVS_D.items():
            if value[0] == v[0]:
                # print(key, k)
                break
        GVL_D[key] = (value, k, v[1])
    return GVL_D


def change_file_name(name,table):
    '''
    读取Spec文件，查找替换req number ，如果查找不到则保持不变
    '''
    if name[0] == 'W':
        require = name[4:9]
    elif name[0] == 'H':
        require = name[3:8]
    if require in table.keys() and len(table[require]) == 3:
        new_head = table[require][-1]+table[require][-2]
        if table[require][0][1] == 'WHAT':
            new_name = new_head + name[11:]
        if table[require][0][1] == 'HOW':
            new_name = new_head + name[10:]
        os.rename(name, new_name)


class execl():
    def __init__(self,name):
        self.app = xw.App(visible=False, add_book=False)
        self.app.display_alerts = False
        self.app.screen_updating = False
        self.wb = self.app.books.open(name)
        self.WB = self.wb.sheets['Test Overview']

    def read_cell(self):
        # HOW = WB.range('B11').value
        return self.WB.range('B10').value

    def write_cell(self, what):
        # HOW = WB.range('B11').value
        self.WB.range('B10').value = what

    def clear(self):
        self.wb.save()
        self.wb.close()
        self.app.quit()

def get_new_requirements(what,table):
    what = what.split(',')
    for i in range(len(what)):
        if len(what[i]) >= 10:
            require = what[i][-7:-2]
            if require in table.keys() and len(table[require]) == 3:
                new_require = table[require][-1] + table[require][-2]
                what[i] = new_require
    new_what = ','.join(what)
    return new_what

def check_folder(folder,table):
    list = os.listdir(folder)
    os.chdir(folder)
    for name in list:
        if os.path.isfile(name):
            spec = execl(name)
            # print(i)
            if re.findall(r'xlsm$',name):
                what = spec.read_cell()
                new_what = get_new_requirements(what,table)
                spec.write_cell(new_what)
                spec.clear()
                change_file_name(name,table)
                print(name)
                print(what,'|',new_what)
                # spec.clear()
        else:
            check_folder(name,table)
    os.chdir('..')

def main():
    # folder = sys.argv[1]
    folder = r'C:\FW_System_Test\JEDI\JEDI_SPEC\WHAT58814 FW Requirement\WHAT69032 Interfaces_Communication\HOW41152 Power module surveillance'
    GVL_D = get_requirements_dict(f_GVL)
    GVS_D = get_requirements_dict(f_GVS)
    table = get_map(GVL_D,GVS_D)
    print(table)
    check_folder(folder,table)


if __name__ == "__main__":
    main()