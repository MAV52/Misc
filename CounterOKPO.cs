public async void CounterOKPO()
{
    var _fd = new SaveFileDialog();
    var filts = new List<FileDialogFilter>();
    filts.Add(new() { Extensions = new string[] { "xlsx" }.ToList() });
    _fd.Filters = filts;
    string? flw = await _fd.ShowAsync(new Window());
    Debug.WriteLine(flw);
    if (flw == null)
    {
        return;
    }
    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
    ExcelPackage xls = new ExcelPackage();
    int row_cur = 1;
    {
        xls = new();
        xls.Workbook.Worksheets.Add("Орг-ии");
        xls.Workbook.Worksheets[0].Cells[1, 1].Value = "Рег.№";
        xls.Workbook.Worksheets[0].Cells[1, 2].Value = "ОКПО";
        xls.Workbook.Worksheets[0].Cells[1, 3].Value = "Краткое наименование";
        xls.Workbook.Worksheets[0].Cells[1, 4].Value = "Номер формы";
        xls.Workbook.Worksheets[0].Cells[1, 5].Value = "Начало ОП";
        xls.Workbook.Worksheets[0].Cells[1, 6].Value = "Окончание ОП";
        xls.Workbook.Worksheets[0].Cells[1, 7].Value = "Коррекция";
        xls.Workbook.Worksheets[0].Cells[1, 8].Value = "Строка";
        xls.Workbook.Worksheets[0].Cells[1, 9].Value = "Код операции";
        xls.Workbook.Worksheets[0].Cells[1, 10].Value = "Дата операции";
        xls.Workbook.Worksheets[0].Cells[1, 11].Value = "ОКПО контрагента";

        using (Orgs_DB DB_MAIN = new(MainWindowViewModel.host_DB_current, MainWindowViewModel.path_DB_current))
        {
            List<string> operations_okpo_self = new() { "10","11","12","15","17","18","41","42","43","46","47","48","53","58","61","62","65","67","68","71","72","73" };
            List<string> operations_okpo_other = new() { "25","27","28","29","35","37","38","39" };
            DB_MAIN.Database.EnsureCreated();
            await DB_MAIN.Database.OpenConnectionAsync();
            List<string[]> orgs = DB_MAIN.DB_Orgs.Where(x => x.viac.ToLower() != "росатом").Select(x => new string[] {
                x.Id.ToString(), x.RegNo_DB, x.Okpo_DB, x.ShortJurLico_DB
            }).ToList();
            for (var i =0;i<orgs.Count;i++)
            {
                while (orgs[i][2].Length < 8) orgs[i][2] = "0"+orgs[i][2];
                while (orgs[i][2].Length > 8 && orgs[i][2].Length < 14) orgs[i][2] = "0"+orgs[i][2];
            }
            List<long> org_Ids = new();
            org_Ids.AddRange(orgs.Select(x => long.Parse(x[0])));

            string form1 = "1. 1";
            string form2 = "1. 2";
            string form3 = "1. 3";
            string form4 = "1. 4";
            List<string[]> forms1 = DB_MAIN.Form_1_1_Archive.Where(x => operations_okpo_self.Contains(x.field_02) || operations_okpo_other.Contains(x.field_02)).Select(
                x => new string[] { x.field_01.ToString(), x.field_02, x.field_03, x.IDF.ToString(), x.field_19, "1.1" }).ToList();
            List<string[]> forms2 = DB_MAIN.Form_1_2_Archive.Where(x => operations_okpo_self.Contains(x.field_02) || operations_okpo_other.Contains(x.field_02)).Select(
                x => new string[] { x.field_01.ToString(), x.field_02, x.field_03, x.IDF.ToString(), x.field_16, "1.2" }).ToList();
            List<string[]> forms3 = DB_MAIN.Form_1_3_Archive.Where(x => operations_okpo_self.Contains(x.field_02) || operations_okpo_other.Contains(x.field_02)).Select(
                x => new string[] { x.field_01.ToString(), x.field_02, x.field_03, x.IDF.ToString(), x.field_17, "1.3" }).ToList();
            List<string[]> forms4 = DB_MAIN.Form_1_4_Archive.Where(x => operations_okpo_self.Contains(x.field_02) || operations_okpo_other.Contains(x.field_02)).Select(
                x => new string[] { x.field_01.ToString(), x.field_02, x.field_03, x.IDF.ToString(), x.field_18, "1.4" }).ToList();
            
            /*
            string form1 = "1. 5";
            string form2 = "1. 6";
            string form3 = "1. 7";
            string form4 = "1. 8";
            List<string[]> forms1 = DB_MAIN.Form_1_5_Archive.Where(x => operations_okpo_self.Contains(x.field_02) || operations_okpo_other.Contains(x.field_02)).Select(
                x => new string[] { x.field_01.ToString(), x.field_02, x.field_03, x.IDF.ToString(), x.field_15, "1.5" }).ToList();
            List<string[]> forms2 = DB_MAIN.Form_1_6_Archive.Where(x => operations_okpo_self.Contains(x.field_02) || operations_okpo_other.Contains(x.field_02)).Select(
                x => new string[] { x.field_01.ToString(), x.field_02, x.field_03, x.IDF.ToString(), x.field_18, "1.6" }).ToList();
            List<string[]> forms3 = DB_MAIN.Form_1_7_Archive.Where(x => operations_okpo_self.Contains(x.field_02) || operations_okpo_other.Contains(x.field_02)).Select(
                x => new string[] { x.field_01.ToString(), x.field_02, x.field_03, x.IDF.ToString(), x.field_17, "1.7" }).ToList();
            List<string[]> forms4 = DB_MAIN.Form_1_8_Archive.Where(x => operations_okpo_self.Contains(x.field_02) || operations_okpo_other.Contains(x.field_02)).Select(
                x => new string[] { x.field_01.ToString(), x.field_02, x.field_03, x.IDF.ToString(), x.field_14, "1.8" }).ToList();
            */

            int rep_cur = 0;
            int rep_lim = DB_MAIN.DB_Reports.Where(x => !x.Annul.Contains("V") && (x.FormNum_DB == form1 || x.FormNum_DB == form2 || x.FormNum_DB == form3 || x.FormNum_DB == form4)).Count();

            List<string[]> forms = new();
            forms.AddRange(forms1);
            forms.AddRange(forms2);
            forms.AddRange(forms3);
            forms.AddRange(forms4);
            for (var i = 0; i < forms.Count; i++)
            {
                while (forms[i][4].Length < 8) forms[i][4] = "0" + forms[i][4];
                while (forms[i][4].Length > 8 && forms[i][4].Length < 14) forms[i][4] = "0" + forms[i][4];
            }
            List<long> form_IDFs = new();
            form_IDFs.AddRange(forms.Select(x => long.Parse(x[3])));
            List<Native_Report> reps = new();
            await DB_MAIN.DB_Reports.Where(x=> !x.Annul.Contains("V") && (x.FormNum_DB == form1 || x.FormNum_DB == form2 || x.FormNum_DB == form3 || x.FormNum_DB == form4)).ForEachAsync(x =>
            {
                rep_cur++;
                Debug.WriteLine($"{rep_cur}/{rep_lim}");
                if (org_Ids.Contains(x.IDF) && form_IDFs.Contains(x.Id))
                {
                    reps.Add(x);
                }
            });
            List<string[]> reports = reps.Select(x=>new string[] {
                x.Id.ToString(),x.IDF.ToString(),misc.FromDate(x.StartPeriod_DB),misc.FromDate(x.EndPeriod_DB),x.CorrectionNumber_DB,"","",""
            }).ToList();
            List<string[]> result = new();
            List<string[]> operations_match = new();
            for (int i=0;i<reports.Count;i++)
            {
                string[]? org_match = orgs.Find(x => x[0] == reports[i][1]);
                if (org_match != null)
                {
                    reports[i][5] = org_match[1];
                    reports[i][6] = org_match[2];
                    reports[i][7] = org_match[3];
                    operations_match = forms.Where(x => x[3] == reports[i][0]).ToList();
                    foreach (string[] operation_match in operations_match)
                    {
                        if ((operations_okpo_self.Contains(operation_match[1]) && operation_match[4].Substring(0,8) != reports[i][6].Substring(0, 8))
                            || (operations_okpo_other.Contains(operation_match[1]) && operation_match[4].Substring(0, 8) == reports[i][6].Substring(0, 8)))
                        {
                            result.Add(new string[]
                            {
                                reports[i][5],
                                reports[i][6].Substring(0,8),
                                reports[i][7],
                                operation_match[5],
                                reports[i][2],
                                reports[i][3],
                                reports[i][4],
                                operation_match[0],
                                operation_match[1],
                                operation_match[2],
                                operation_match[4],
                                reports[i][1],
                                reports[i][0]
                            });
                        }
                    }
                }
            }
            //remove old corrections
            for (int i = 0; i < result.Count; i++)
            {
                string cor_max = result.Where(x =>
                    x[0] == result[i][0]
                    && x[1] == result[i][1]
                    && x[2] == result[i][2]
                    && x[3] == result[i][3]
                    && x[4] == result[i][4]
                    && x[5] == result[i][5]
                    && x[11] == result[i][11]
                    && x[12] == result[i][12]
                ).Max(x => int.Parse(x[6])).ToString();
                result[i][6] = "PH";
                result.RemoveAll(x =>
                    x[0] == result[i][0]
                    && x[1] == result[i][1]
                    && x[2] == result[i][2]
                    && x[3] == result[i][3]
                    && x[4] == result[i][4]
                    && x[5] == result[i][5]
                    && x[11] == result[i][11]
                    && x[12] == result[i][12]
                    && x[6] != "PH");
                result[i][6] = cor_max;
            }
            List<Native_Report> repps = new();
            repps = DB_MAIN.DB_Reports.Where(x=> x.FormNum_DB == form1 || x.FormNum_DB == form2 || x.FormNum_DB == form3 || x.FormNum_DB == form4).ToList();
            DateTime trytime;
            int tryint;
            int ii = 0;
            for (ii = 0; ii < result.Count; ii++)
            {
                List<string[]> corrs = repps.Where(x =>
                x.IDF == long.Parse(result[ii][11])
                && x.StartPeriod_DB == DateTime.Parse(DateTime.TryParse(result[ii][04], out trytime) ? result[ii][04] : "01.01.1970")
                && x.EndPeriod_DB == DateTime.Parse(DateTime.TryParse(result[ii][05], out trytime) ? result[ii][05] : "01.01.1970")
                && x.FormNum_DB.Replace(" ","").Equals(result[ii][03].Replace(" ", ""))
                && int.Parse(int.TryParse(x.CorrectionNumber_DB,out tryint)? x.CorrectionNumber_DB : "0") > int.Parse(result[ii][06])).Select(x => new string[] {
                    x.Id.ToString(),
                    x.FormNum_DB,
                    x.CorrectionNumber_DB.ToString()
                }).ToList();
                if (corrs.Count > 0)
                {
                    string[] corr = corrs.First(x => int.Parse(x[2]) == corrs.Max(y => int.TryParse(y[2],out tryint)?tryint:0));
                    List<string[]> opers = new();
                    if (corr[1] == form1)
                    {
                        opers = forms1.Where(x => x[3] == corr[0] && x[0] == result[ii][07]).Select(x => new string[] { x[1], x[4] }).ToList();
                    }
                    else if (corr[1] == form2)
                    {
                        opers = forms2.Where(x => x[3] == corr[0] && x[0] == result[ii][07]).Select(x => new string[] { x[1], x[4] }).ToList();
                    }
                    else if (corr[1] == form3)
                    {
                        opers = forms3.Where(x => x[3] == corr[0] && x[0] == result[ii][07]).Select(x => new string[] { x[1], x[4] }).ToList();
                    }
                    else if (corr[1] == form4)
                    {
                        opers = forms4.Where(x => x[3] == corr[0] && x[0] == result[ii][07]).Select(x => new string[] { x[1], x[4] }).ToList();
                    }
                    if (opers.Count > 0)
                    {
                        if ((operations_okpo_self.Contains(opers.First()[0]) && opers.First()[1].Substring(0, 8) == result[ii][1].Substring(0, 8))
                            || (operations_okpo_other.Contains(opers.First()[0]) && opers.First()[1].Substring(0, 8) != result[ii][1].Substring(0, 8)))
                        {
                            result.RemoveAt(ii);
                            ii--;
                        }
                    }
                }
            }
            row_cur = 1;
            foreach (string[] res in result)
            {
                row_cur++;
                xls.Workbook.Worksheets[0].Cells[row_cur, 1].Value = res[0];
                xls.Workbook.Worksheets[0].Cells[row_cur, 2].Value = res[1];
                xls.Workbook.Worksheets[0].Cells[row_cur, 3].Value = res[2];
                xls.Workbook.Worksheets[0].Cells[row_cur, 4].Value = res[3];
                xls.Workbook.Worksheets[0].Cells[row_cur, 5].Value = res[4];
                xls.Workbook.Worksheets[0].Cells[row_cur, 6].Value = res[5];
                xls.Workbook.Worksheets[0].Cells[row_cur, 7].Value = res[6];
                xls.Workbook.Worksheets[0].Cells[row_cur, 8].Value = res[7];
                xls.Workbook.Worksheets[0].Cells[row_cur, 9].Value = res[8];
                xls.Workbook.Worksheets[0].Cells[row_cur, 10].Value = res[9];
                xls.Workbook.Worksheets[0].Cells[row_cur, 11].Value = res[10];
            }
            await DB_MAIN.Database.CloseConnectionAsync();
        }
        xls.SaveAs(flw);
    }
}