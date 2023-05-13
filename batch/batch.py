import pandas as pd
from datetime import datetime
import win32com.client as win32
import os
import pythoncom
from MSSQLDB_connect import MSSQLDB
import json
import logging

class Auto_excel():
    
    def __init__(self):
        now = datetime.now()
        log_filename = 'auto_excel_{}.log'.format(now.strftime('%Y-%m-%d'))
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(asctime)s:%(levelname)s:%(message)s')
        log_dir = '.\\var\\log'
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        file_handler = logging.FileHandler('./var/log/' + log_filename)
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)
        
    def generate_report(self):
        try:
            cfg_path = r"C:\Users\gary1\OneDrive\文件\GitHub\Wiwynn_excel_automation_python_project\config.json"
            with open(cfg_path, 'r', encoding='utf-8-sig') as f:
                cfg = json.load(f)
            # databse connection information
            db_cfg = cfg['db_connect']
            db_cfg['creator'] = __import__(db_cfg['creator'])
            db = MSSQLDB(db_cfg)
            df = db.read_table('DataCoSupplyChainDataset')
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            df['order date (DateOrders)'] = df['order date (DateOrders)'].apply(lambda x: datetime.strptime(x, '%m/%d/%Y %H:%M').strftime('%Y/%m/%d'))
            df['order date (DateOrders)'] = pd.to_datetime(df['order date (DateOrders)'])
            # 篩選出第一季的資料
            q1_data = df[(df['order date (DateOrders)'].dt.quarter == 1) & (df['order date (DateOrders)'].dt.year == 2017)]
            product = q1_data.groupby('Product Name')[['Sales','Order Item Quantity']].sum().reset_index()
            a = q1_data.groupby('Product Name')['Category Name'].unique().reset_index()
            product = pd.merge(product,a)
            product['Category Name'] = product['Category Name'].apply(lambda x : x[0])
            product_Category = q1_data.groupby('Category Name')[['Order Item Quantity','Sales']].sum().reset_index()
            # 新增分頁前先刪除同名分頁
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.DisplayAlerts = False
            workbook = excel.Workbooks.Add()
            try:
                worksheet = workbook.Worksheets('Q1 Product Sales')
                worksheet.Delete()
            except:
                pass
            try:
                worksheet_1 = workbook.Worksheets('Q1 Category Sales')
                worksheet_1.Delete()
            except:
                pass
            # 新增分頁
            worksheet = workbook.Worksheets.Add()
            worksheet.Name = 'Q1 Product Sales'
            for i in range(len(product.columns)):
                worksheet.Cells(2, i+1).Value = product.columns[i]
            for i in range(len(product)):
                for j in range(len(product.columns)):
                    worksheet.Cells(i+3, j+1).Value = product.iloc[i,j]
            worksheet.Cells.EntireColumn.AutoFit()
            # 設置標題
            title_range = worksheet.Range("A1", "E1")
            title_range.Merge()
            title_range.Value = "Q1 Product Sales Report"
            title_range.Font.Size = 16
            title_range.Font.Bold = True
            title_range.HorizontalAlignment = win32.constants.xlCenter
            worksheet_1 = workbook.Worksheets.Add()
            worksheet_1.Name = 'Q1 Category Sales'
            for i in range(len(product_Category.columns)):
                worksheet_1.Cells(2, i+1).Value = product_Category.columns[i]
            for i in range(len(product_Category)):
                for j in range(len(product_Category.columns)):
                    worksheet_1.Cells(i+3, j+1).Value = product_Category.iloc[i,j]
            worksheet_1.Cells.EntireColumn.AutoFit()
            # 設置標題

            title_range_1 = worksheet_1.Range("A1", "C1")
            title_range_1.Merge()
            title_range_1.Value = "Q1 Category Sales Report"
            title_range_1.Font.Size = 16
            title_range_1.Font.Bold = True
            title_range_1.HorizontalAlignment = win32.constants.xlCenter
            # 新增圖表
            chart_range = worksheet.Range("A3", "B" + str(len(product) + 2))
            chart = worksheet.Shapes.AddChart2(251, 4, 600, 100).Chart
            chart.SetSourceData(chart_range)
            chart.ChartTitle.Text = "Sales by Product Name"
            chart.HasTitle = True
            chart.ChartType = win32.constants.xlColumnClustered
            chart.Axes(win32.constants.xlValue).HasTitle = True
            chart.Axes(win32.constants.xlValue).AxisTitle.Text = "Sales"
            chart.Axes(win32.constants.xlCategory).HasTitle = True
            chart.Axes(win32.constants.xlCategory).AxisTitle.Text = "Product Name"
            chart.Top = worksheet.Range("H3").Top
            chart.Left = worksheet.Range("H3").Left
            chart.Height = 250
            chart.Width = 400

            chart1_range = worksheet_1.Range("A3:A26,C3:C26")
            chart1 = worksheet_1.Shapes.AddChart2(251, 4, 300, 100).Chart
            chart1.SetSourceData(chart1_range)
            chart1.ChartTitle.Text = "Sales by Category Name"
            chart1.HasTitle = True
            chart1.ChartType = win32.constants.xlPie
            chart1.Top = worksheet_1.Range("A12").Top
            chart1.Left = worksheet_1.Range("A12").Left
            chart1.Height = 600
            chart1.Width = 1200

            labels = worksheet_1.Range("A3", "A26").Value
            values = worksheet_1.Range("C3", "C26").Value
            current_directory = os.getcwd()
            now = datetime.now()
            filename = "report_" + now.strftime("%Y-%m-%d_%H-%M") + ".xlsx"
            filepath = os.path.join(current_directory, filename).replace("/", "\\")
            workbook.SaveAs(filepath)
            workbook.Close()
            excel.Quit()
            excel.DisplayAlerts = True
            self.logger.info('Report generated successfully.')
        except Exception as e:
            self.logger.error(str(e))