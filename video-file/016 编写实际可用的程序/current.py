if __name__ == '__main__':

    wb = openpyxl.Workbook()
    ws = wb.active
    # 判断sys.argv的长度

    if len(sys.argv)!=1 and len(sys.argv)!=2:
        print("参数错误")

    elif len(sys.argv) == 1:
        sales_list(ws, 20).parent.save("new.xlsx")
    else:
        # 还可以加一些判断, 比如文件名不合法
        try:
            sales_list(ws, 20).parent.save(sys.argv[1])
        except:
            print("文件名不合法.")
