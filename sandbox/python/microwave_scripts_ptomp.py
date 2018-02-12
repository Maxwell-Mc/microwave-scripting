


def get_ciq_cascade(args):
    from openpyxl import load_workbook
    import openpyxl
    data = {}
    cascade = []
    status = False
    try:
        if args.ciq:
            wb = load_workbook(args.ciq)
            # Loop through sheets
            for ws in wb.worksheets:
                data[ws.title] = {}
                consecutive_col_none = 0
                consecutive_row_none = 0
                for col in range(1, ws.max_column + 1):
                    if ws[openpyxl.utils.get_column_letter(col) + '1' ].value == None:
                        consecutive_col_none = consecutive_col_none + 1
                    else:
                        consecutive_col_none = 0
                for row in range(1, ws.max_row + 1):
                    if ws[openpyxl.utils.get_column_letter(1) + str(row) ].value == None:
                        consecutive_row_none = consecutive_row_none + 1
                    else:
                        consecutive_row_none = 0
                for row in range(1, ws.max_row - consecutive_row_none):
                    data[ws.title][row]= {}
                    for col in range(1, ws.max_column - consecutive_col_none):
                        data[ws.title][row][ws[openpyxl.utils.get_column_letter(col) + '1'].value] = ws[openpyxl.utils.get_column_letter(col) + str(row) ].value

        for key, value in data['CIQs'].iteritems():
            if 'Site ID' in value.keys():
                if args.cascade in str(value['Site ID']):
                    cascade.append(value)
        
        status = True
    except:
        status = False
    
    return status, cascade

def get_sheet_count(ws, h, data):
    import openpyxl
    consecutive_col_none = 0
    consecutive_row_none = 0
    # data = {}
    for col in range(1, ws.max_column + 1):
        if ws[openpyxl.utils.get_column_letter(col) + str(h) ].value == None:
            consecutive_col_none = consecutive_col_none + 1
        else:
            consecutive_col_none = 0
    for row in range(h, ws.max_row + 1):
        if ws[openpyxl.utils.get_column_letter(1) + str(row) ].value == None:
            consecutive_row_none = consecutive_row_none + 1
        else:
            consecutive_row_none = 0
    for row in range(1, ws.max_row - consecutive_row_none):
                    data[ws.title][row]= {}
                    for col in range(1, ws.max_column - consecutive_col_none):
                        data[ws.title][row][ws[openpyxl.utils.get_column_letter(col) + str(h)].value] = ws[openpyxl.utils.get_column_letter(col) + str(row) ].value
    return data[ws.title]

def get_ip_plan_cascade(args):
    from openpyxl import load_workbook
    import openpyxl
    data = {}
    cascade = {}
    status = False
    try:
        if args.ip_plan:
            wb = load_workbook(args.ip_plan)
            # Loop through sheets
            for ws in wb.worksheets:
                data[ws.title] = {}
                if '3G IP Plan' in ws.title:
                    data[ws.title] = get_sheet_count(ws, 3, data)
                elif '4G IP Plan' in ws.title:
                    data[ws.title] = get_sheet_count(ws, 2, data)
                else:
                    data[ws.title] = get_sheet_count(ws, 1, data)

        ip_plan_3g = []
        for key, value in data['3G IP Plan'].iteritems():
            field = "Cascade_ID\ncell site-id"
            if field in value.keys():
                if args.cascade in str(value[field]):
                    ip_plan_3g.append(value)
        cascade['3G IP Plan'] = ip_plan_3g
        ip_plan_4g = []
        for key, value in data['4G IP Plan'].iteritems():
            field = "Cascade_ID\ncell site-id"
            if field in value.keys():
                if args.cascade in str(value[field]):
                    ip_plan_4g.append(value)
        cascade['4G IP Plan'] = ip_plan_4g
    except:
        pass
    return True, cascade

def main():
    import argparse
    from pprint import pprint

    data ={}

    parser = argparse.ArgumentParser(description='Process to load excel ciq documents')
    parser.add_argument('-q','--ciq', dest='ciq', help='load to FILE', metavar='FILE')
    parser.add_argument('-p','--ip_plan', dest='ip_plan', help='load to FILE', metavar='FILE')
    parser.add_argument('-c','--cascade', dest='cascade', help='search on the cascade', metavar='cascade')

    args = parser.parse_args()

    ciq_status, cascade_ciq = get_ciq_cascade(args)
    if ciq_status:
        data['ciq'] = cascade_ciq
    else:
        data['ciq'] ='Error: 02'
    
    plan_status, cascade_plan = get_ip_plan_cascade(args)
    if plan_status:
        data['ip_plan'] = cascade_plan

    pprint(data)


if __name__ == '__main__':
    main()