import pythoncom, win32com.client
DATA=[["大区","城市","销售额"],["华东","上海",320],["华东","杭州",210],["华南","广州",280],["华南","深圳",350]]
pythoncom.CoInitialize()
xl=win32com.client.Dispatch("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
bk=None
try:
    bk=xl.Workbooks.Add(); ws=bk.Worksheets(1); ws.Range("A1:C5").Value=DATA
    for label, fn in [
        ("K. ChartWizard 设数据源", lambda c: c.ChartWizard(ws.Range("A1:C5"))),
        ("L. 显式 Series.Values+层级", None),
    ]:
        sh=ws.Shapes.AddChart2(-1,116,10,10,300,200); ch=sh.Chart
        if fn:
            try: fn(ch)
            except Exception as e: print(f"  {label}: wizard failed {str(e)[:50]}")
        else:
            try:
                s=ch.SeriesCollection().NewSeries(); s.Values=ws.Range("C2:C5"); s.XValues=ws.Range("A2:B5")
            except Exception as e: print(f"  {label}: {str(e)[:50]}")
        try:
            t=int(ch.ChartType); print(f"  {label}: type={t} {'OK' if t==116 else '回落'}")
        except Exception as e: print(f"  {label}: read failed")
        sh.Delete()
    # 也试一下独立图表页方式
    try:
        chs=bk.Charts.Add(); chs.ChartType=116; print(f"  M. Charts.Add + ChartType=116: type={int(chs.ChartType)}")
    except Exception as e:
        print(f"  M. Charts.Add: failed {str(e)[:60]}")
finally:
    if bk is not None: bk.Close(SaveChanges=False)
    xl.Quit(); pythoncom.CoUninitialize()
