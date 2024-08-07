  public static void InsertIDsListIntoExcel(string filePath, List<int> data)
        {
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Sheets[1];
            int rowindex = 2;
            foreach (int eachdata in data)
            {
                worksheet.Cells[rowindex++, 2] = eachdata;
            }


            workbook.Save();



            workbook.Close(false);
            excelApp.Quit();
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        public static void CompareIDsInExcel(string filePath, string colName, string sheetName, List<int> ListOfIDS, List<int> CompareListIDS)
        {
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Sheets[sheetName];
            int rowindex = 1;
            Excel.Range Usedrange = worksheet.UsedRange;
            int colCount = Usedrange.Columns.Count;
            int rowCount = Usedrange.Rows.Count;
            List<int> FailIDs = new List<int>();
            int FailIDIndex = 1;
            for (int i = 0; i < CompareListIDS.Count; i++)
            {
                FailIDIndex = 1;
                for (int j = 0; j < ListOfIDS.Count; j++)
                {
                    FailIDIndex++;
                    if (CompareListIDS[i] == ListOfIDS[j])
                    {
                        FailIDs.Add(FailIDIndex);
                        break;
                    }
                }
            }

            worksheet.Cells[1, colCount + 1] = colName;
            for (int i = 0; i < FailIDs.Count; i++)
            {
                worksheet.Cells[FailIDs[i], colCount + 1]="FAIL";

            }
            workbook.Save();



            workbook.Close(false);
            excelApp.Quit();
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        public static string GetAlphabetWithNumber(int colNum)
        {
            string resultName = "";
            resultName = ((char)(colNum + 64)).ToString();
            return resultName;

        }
        public static void ContTheFailsInRow(string filePath)
        {
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Sheets[1];
            int rowindex = 1;
            Excel.Range Usedrange = worksheet.UsedRange;
            int colCount = Usedrange.Columns.Count;
            int rowCount = Usedrange.Rows.Count;
            string lastCol=GetAlphabetWithNumber(colCount);
            Excel.Range formulaRange = (Excel.Range)worksheet.Range["A2"+ ":A" + rowCount];

            string valueToCount = "FAIL";

            string formrange = string.Format("B2:{0}2", lastCol);
            string countifFormula = $"=COUNTIF({formrange},\"{valueToCount}\")";
          
            formulaRange.Formula = countifFormula;
            workbook.Save();




            workbook.Close(false);
            excelApp.Quit();
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
