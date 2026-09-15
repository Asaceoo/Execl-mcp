import pythoncom, win32com.client
DATA=[["Region","Q1","Q2"],["华东",100,130],["华南",200,170],["华北",300,240],["西南",180,260],["东北",90,150],["西北",140,110]]
CASES=[("Treemap",117),("Sunburst",116),("Histogram",118),("Pareto",122),("BoxWhisker",121),("Waterfall",119),("Funnel",123),("RegionMap",140),("ColumnLineCombo",120)]
pythoncom.CoInitialize()
xl=win32com.client.Dispatch("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
try:
    bk=xl.Workbooks.Add(); ws=bk.Worksheets(1); ws.Range("A1:C7").Value=DATA
    print(f"{'type':<17}{'E: AddChart2+NewLayout':<26}{'F: 手动 NewSeries':<26}{'G: 只读ChartType'}")
    for name,code in CASES:
        res=[]
        # E
        try:
            sh=ws.Shapes.AddChart2(-1,code,10,10,300,200,True); ch=sh.Chart
            ch.SetSourceData(ws.Range("A1:C7"),2); a=int(ch.ChartType); sh.Delete()
            res.append("OK" if a==code else f"回落{a}")
        except Exception as e: res.append("FAIL")
        # F
        try:
            sh=ws.Shapes.AddChart2(-1,code,10,10,300,200); ch=sh.Chart
            s=ch.SeriesCollection().NewSeries(); s.Values=ws.Range("B2:B7"); s.XValues=ws.Range("A2:A7")
            a=int(ch.ChartType); sh.Delete(); res.append("OK" if a==code else f"回落{a}")
        except Exception as e: res.append("FAIL")
        # G
        try:
            sh=ws.Shapes.AddChart2(-1,code,10,10,300,200); a=int(sh.Chart.ChartType); sh.Delete()
            res.append(f"type={a}")
        except Exception as e: res.append("FAIL")
        print(f"{name:<17}{res[0]:<26}{res[1]:<26}{res[2]}")
finally:
    bk.Close(SaveChanges=False); xl.Quit(); pythoncom.CoUninitialize()
