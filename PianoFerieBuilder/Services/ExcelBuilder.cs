using OfficeOpenXml;
using OfficeOpenXml.Table;
using PianoFerieBuilder.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PianoFerieBuilder.Services
{
    class ExcelBuilder
    {
        private static int HEADER = 1;
        private static string WORKSHEET_NAME = "Calendar";
        private static string TABLE_NAME = "Calendar";

        public int CreateFile(string path, List<CalendarDay> calendar)
        {
            FileInfo fi = new FileInfo(path);

            if (!string.Equals(fi.Extension, ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("Invalid file extension: use .xlsx");
            }

            using (ExcelPackage package = new ExcelPackage(fi))
            {
                var workBook = package.Workbook;

                int toalRosw = calendar.Count + HEADER;

                ExcelWorksheet workSheet = workBook.Worksheets.FirstOrDefault(v => v.Name == WORKSHEET_NAME);

                if (workSheet == null)
                {
                    workSheet = workBook.Worksheets.Add(WORKSHEET_NAME);
                }

                ExcelTable newTable = workSheet.Tables.FirstOrDefault(v => v.Name == TABLE_NAME);

                if (newTable == null)
                {
                    string letter = ExcelCellBase.GetAddressCol(8);
                    newTable = workSheet.Tables.Add(new ExcelAddressBase($"A1:{letter}{toalRosw}"), TABLE_NAME);
                }

                PopulateWorksheet(workSheet, newTable, calendar);

                package.Save();
            }

            return 0;
        }


        private void PopulateWorksheet(ExcelWorksheet workSheet, ExcelTable table, List<CalendarDay> data)
        {
            workSheet.Cells["A1"].Value = "Giorno";
            workSheet.Cells["B1"].Value = "Lavorato";
            workSheet.Cells["C1"].Value = "Ferie";
            workSheet.Cells["D1"].Value = "ROL";
            workSheet.Cells["E1"].Value = "Ex-Festività";
            workSheet.Cells["F1"].Value = "Rip Compensazione";
            workSheet.Cells["G1"].Value = "Tot";
            workSheet.Cells["H1"].Value = "Note";

            table.Columns[6].CalculatedColumnFormula = "SUM($B2:$E2)";

            for (int i = 0; i < data.Count; i++)
            {
                workSheet.Cells[$"A{i + 2}"].Formula = $"DATE({data[i].Date.Year},{data[i].Date.Month},{data[i].Date.Day})";

                int oreLavoro = data[i].IsWeekend || data[i].IsHoliday ? 0 : 8;

                workSheet.Cells[$"B{i + 2}"].Value = oreLavoro;
                workSheet.Cells[$"C{i + 2}"].Value = 0;
                workSheet.Cells[$"D{i + 2}"].Value = 0;
                workSheet.Cells[$"E{i + 2}"].Value = 0;
                workSheet.Cells[$"F{i + 2}"].Value = 0;

                if (data[i].IsWeekend && !data[i].IsHoliday)
                {
                    workSheet.Cells[$"A{i + 2}:H{i + 2}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[$"A{i + 2}:H{i + 2}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSkyBlue);
                }

                if (!data[i].IsWeekend && data[i].IsHoliday)
                {
                    workSheet.Cells[$"A{i + 2}:H{i + 2}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[$"A{i + 2}:H{i + 2}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightCoral);
                }

                if (data[i].IsWeekend && data[i].IsHoliday)
                {
                    workSheet.Cells[$"A{i + 2}:H{i + 2}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    workSheet.Cells[$"A{i + 2}:H{i + 2}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkRed);
                    workSheet.Cells[$"A{i + 2}:H{i + 2}"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                }
            }

            //// Out of table
            workSheet.Cells["J1"].Value = "Mese Corrente";
            workSheet.Cells["J2"].Value = "Ore Giorno";
            workSheet.Cells["J3"].Value = "Ferie Mese";
            workSheet.Cells["J4"].Value = "Ex-Festività";
            workSheet.Cells["J5"].Value = "ROL";
            workSheet.Cells["J6"].Value = "Giorni lavorativi";

            workSheet.Cells["K1"].Value = 1;
            workSheet.Cells["K2"].Value = 8;
            workSheet.Cells["K3"].Formula = "L3/12";
            workSheet.Cells["K4"].Formula = "L4/12";
            workSheet.Cells["K5"].Formula = "L5/12";
            workSheet.Cells["K6"].Value = data.Where(v => !v.IsHoliday && !v.IsWeekend).Count();

            workSheet.Cells["L3"].Value = 176;
            workSheet.Cells["L4"].Value = 32;
            workSheet.Cells["L5"].Value = 72;
            workSheet.Cells["L6"].Formula = "SUM(L3:L5)/$K$2";

            workSheet.Cells["M1"].Value = "Limite";
            workSheet.Cells["N1"].Formula = $"DATE({data[0].Date.Year},K1,1)";

            workSheet.Cells["R12"].Value = "Proiezione fine anno";

            workSheet.Cells["K13"].Value = "Residue";
            workSheet.Cells["L13"].Value = "Godute";
            workSheet.Cells["M13"].Value = "In Busta";
            workSheet.Cells["N13"].Value = "Maturate";
            workSheet.Cells["O13"].Value = "Anno Prec.";
            workSheet.Cells["P13"].Value = "-";
            workSheet.Cells["R13"].Value = "Residue";
            workSheet.Cells["S13"].Value = "Godute";

            workSheet.Cells["J14"].Value = "Ferie";
            workSheet.Cells["J15"].Value = "Ex-Festività";
            workSheet.Cells["J16"].Value = "ROL";
            workSheet.Cells["J18"].Value = "Rip Comp.";
            workSheet.Cells["J19"].Value = "TOT ORE";
            workSheet.Cells["J20"].Value = "TOT GIORNI";

            workSheet.Cells["K14"].Formula = "O14+N14-L14";
            workSheet.Cells["K15"].Formula = "O15+N15-L15";
            workSheet.Cells["K16"].Formula = "O16+N16-L16";
            workSheet.Cells["K17"].Formula = "SUM(K15:K16)";
            workSheet.Cells["K18"].Formula = "O18+N18-L18";
            workSheet.Cells["K19"].Formula = "SUM(K$14:K$16)";
            workSheet.Cells["K20"].Formula = "K19/$K$2";
            
            workSheet.Cells["L14"].Formula = "SUMIFS(Calendar[[Ferie]:[Ferie]],Calendar[[Giorno]:[Giorno]],\"<\"&N$1)+M14";
            workSheet.Cells["L15"].Formula = "SUMIFS(Calendar[[Ex-Festività]:[Ex-Festività]],Calendar[[Giorno]:[Giorno]],\"<\"&N$1)+M15";
            workSheet.Cells["L16"].Formula = "SUMIFS(Calendar[[ROL]:[ROL]],Calendar[[Giorno]:[Giorno]],\"<\"&N$1)+M16";
            workSheet.Cells["L17"].Formula = "SUM(L15:L16)";
            workSheet.Cells["L18"].Formula = "SUMIFS(Calendar[[Rip Compensazione]:[Rip Compensazione]],Calendar[[Giorno]:[Giorno]],\"<\"&N$1)+M16";
            workSheet.Cells["L19"].Formula = "SUM(L$14:L$16)+L18";
            workSheet.Cells["L20"].Formula = "L19/$K$2";

            workSheet.Cells["M14"].Value = 0;
            workSheet.Cells["M15"].Value = 0;
            workSheet.Cells["M16"].Value = 0;
            workSheet.Cells["M17"].Value = 0;
            workSheet.Cells["M18"].Value = 0;
            workSheet.Cells["M19"].Formula = "SUM(M$14:M$16)+M18";
            workSheet.Cells["M20"].Formula = "M19/$K$2";

            workSheet.Cells["N14"].Formula = "$K$1*$K3";
            workSheet.Cells["N15"].Formula = "$K$1*$K4";
            workSheet.Cells["N16"].Formula = "$K$1*$K5";
            workSheet.Cells["N17"].Formula = "SUM(N15:N16)";
            workSheet.Cells["N18"].Value = 0;
            workSheet.Cells["N19"].Formula = "SUM(N$14:N$16)+N18";
            workSheet.Cells["N20"].Formula = "N19/$K$2";

            workSheet.Cells["O14"].Value = 0;
            workSheet.Cells["O15"].Value = 0;
            workSheet.Cells["O16"].Value = 0;
            workSheet.Cells["O17"].Value = 0;
            workSheet.Cells["O18"].Value = 0;
            workSheet.Cells["O19"].Formula = "SUM(O$14:O$16)+O18";
            workSheet.Cells["O20"].Formula = "O19/$K$2";

            workSheet.Cells["P13"].Value = "-";
            workSheet.Cells["P14"].Value = "-";
            workSheet.Cells["P15"].Value = "-";
            workSheet.Cells["P16"].Value = "-";
            workSheet.Cells["P17"].Value = "-";
            workSheet.Cells["P18"].Value = "-";
            workSheet.Cells["P19"].Value = "-";
            workSheet.Cells["P20"].Value = "-";

            workSheet.Cells["Q14"].Value = "Ferie";
            workSheet.Cells["Q15"].Value = "Ex-Festività";
            workSheet.Cells["Q16"].Value = "ROL";
            workSheet.Cells["Q18"].Value = "Rip Comp.";
            workSheet.Cells["Q19"].Value = "TOT ORE";
            workSheet.Cells["Q20"].Value = "TOT GIORNI";

            workSheet.Cells["R14"].Formula = "$L3+$O14-$S14";
            workSheet.Cells["R15"].Formula = "$L4+$O15-$S15";
            workSheet.Cells["R16"].Formula = "$L5+$O16-$S16";
            workSheet.Cells["R17"].Formula = "SUM(R15:R16)";
            workSheet.Cells["R18"].Formula = "$L7+$O18-$S18";
            workSheet.Cells["R19"].Formula = "SUM(R$14:R$16)+R18";
            workSheet.Cells["R20"].Formula = "R19/$K$2";

            workSheet.Cells["S14"].Formula = "SUM(Calendar[[Ferie]:[Ferie]])+M14";
            workSheet.Cells["S15"].Formula = "SUM(Calendar[[Ex-Festività]:[Ex-Festività]])+M15";
            workSheet.Cells["S16"].Formula = "SUM(Calendar[[ROL]:[ROL]])+M16";
            workSheet.Cells["S17"].Formula = "SUM(S15:S16)";
            workSheet.Cells["S18"].Formula = "SUM(Calendar[[Rip Compensazione]:[Rip Compensazione]])+M18";
            workSheet.Cells["S19"].Formula = "SUM(S$14:S$16)+S18";
            workSheet.Cells["S20"].Formula = "S19/$K$2";
        }
    }
}
