using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace EducationalProcces
{
    public class Excel
    {
        /// <summary>
        /// Объявление приложения
        /// Отображение Excel
        /// Отключение отображения окон с сообщениями
        /// Получение первого листа документа (счет начинается с 1)
        /// Название листа (вкладки снизу)
        /// Захватываемый диапазон ячеек 
        /// Заполнение ячеек данными из базы данных
        /// Выравнивание ячеек по центру
        /// Объединение ячеек под определенную группу
        /// </summary>
        public async void excel()
        {
            Microsoft.Office.Interop.Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
            ex.Visible = true;
            ex.Workbooks.Open(@"C:\EducationalProcces\EducationalProcces\Расписание.xlsx", Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            ex.DisplayAlerts = false;
            Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)ex.Sheets[1];
            sheet.Name = "КУГ";
            Microsoft.Office.Interop.Excel.Range range = sheet.get_Range("A1", "DC300");
            
            Group group = new Group();
            ResponseModel<List<Group>> responsGroup = await APIHelper.GetDataAsync<List<Group>>("\"\"", "\"\"", "\"\"", typeof(Group));

            if (responsGroup.StatusCode == 201)
            {
                try
                {
                    range.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    int a = 0;
                    for (int i = 14; i <= 115; i++)
                    {
                        for (int j = 1; j < 2; j++)
                        {
                            sheet.Cells[i, j] = String.Format(responsGroup.Data[a].Name_Group, i, j);
                            if (responsGroup.Data[a].Name_Group == "Э-1-21" 
                                || responsGroup.Data[a].Name_Group == "Э-2-21")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();

                                for (int dot = 84; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if(responsGroup.Data[a].Name_Group == "Э-1-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int ochi = 40; ochi < 45; ochi++)
                                    sheet.Cells[i, ochi] = String.Format("ОЦИ", i, ochi);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 44]].Merge();

                                for (int two = 45; two < 46; two++)
                                    sheet.Cells[i, two] = String.Format("2", i, two);
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();

                                for (int ps = 76; ps < 79; ps++)
                                    sheet.Cells[i, ps] = String.Format("ПЭС", i, ps);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 78]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int four = 87; four < 88; four++)
                                    sheet.Cells[i, four] = String.Format("4", i, four);

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if(responsGroup.Data[a].Name_Group == "Э-2-20, Э-11-21")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();

                                for (int ochi = 46; ochi < 51; ochi++)
                                    sheet.Cells[i, ochi] = String.Format("ОЦИ", i, ochi);
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 50]].Merge();

                                for (int two = 51; two < 52; two++)
                                    sheet.Cells[i, two] = String.Format("2", i, two);
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();

                                for (int ps = 80; ps < 83; ps++)
                                    sheet.Cells[i, ps] = String.Format("ПЭС", i, ps);
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 82]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int four = 87; four < 88; four++)
                                    sheet.Cells[i, four] = String.Format("4", i, four);

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if(responsGroup.Data[a].Name_Group == "Э-1-19, Э-11/1-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int rpms = 40; rpms < 46; rpms++)
                                    sheet.Cells[i, rpms] = String.Format("РПМС", i, rpms);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();

                                for (int aos = 76; aos < 78; aos++)
                                    sheet.Cells[i, aos] = String.Format("АОС", i, aos);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();

                                for (int dot = 82; dot < 84; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();

                                for (int pp = 84; pp < 90; pp++)
                                    sheet.Cells[i, pp] = String.Format("ПП.02.01", i, pp);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 89]].Merge();

                                for (int equally1 = 90; equally1 < 108; equally1++)
                                    sheet.Cells[i, equally1] = String.Format("=", i, equally1);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if(responsGroup.Data[a].Name_Group == "Э-2-19, Э-11/2-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();

                                for (int rpms = 46; rpms < 50; rpms++)
                                    sheet.Cells[i, rpms] = String.Format("РПМС", i, rpms);
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();

                                for (int aos = 80; aos < 82; aos++)
                                    sheet.Cells[i, aos] = String.Format("АОС", i, aos);
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();

                                for (int dot = 82; dot < 84; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();

                                for (int pp = 84; pp < 90; pp++)
                                    sheet.Cells[i, pp] = String.Format("ПП.02.01", i, pp);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 89]].Merge();

                                for (int equally1 = 90; equally1 < 108; equally1++)
                                    sheet.Cells[i, equally1] = String.Format("=", i, equally1);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if(responsGroup.Data[a].Name_Group == "Э-1-18")
                            {
                                for (int pp2 = 2; pp2 < 22; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();

                                for (int toir = 22; toir < 25; toir++)
                                    sheet.Cells[i, toir] = String.Format("ТОиР", i, toir);
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 24]].Merge();

                                for (int pp2 = 25; pp2 < 35; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int two = 35; two < 36; two++)
                                    sheet.Cells[i, two] = String.Format("2", i, two);

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int pp2 = 40; pp2 < 66; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();

                                for (int dot = 66; dot < 68; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();

                                for (int pdp = 68; pdp < 76; pdp++)
                                    sheet.Cells[i, pdp] = String.Format("ПДП", i, pdp);
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 75]].Merge();

                                for (int pvkr = 76; pvkr < 84; pvkr++)
                                    sheet.Cells[i, pvkr] = String.Format("ПВКР", i, pvkr);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 83]].Merge();

                                for (int zvkr = 84; zvkr < 88; zvkr++)
                                    sheet.Cells[i, zvkr] = String.Format("ЗВКР", i, zvkr);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if(responsGroup.Data[a].Name_Group == "Э-2-18, Э-11-19")
                            {
                                for (int pp2 = 2; pp2 < 25; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();

                                for (int toir = 25; toir < 28; toir++)
                                    sheet.Cells[i, toir] = String.Format("ТОиР", i, toir);
                                sheet.Range[sheet.Cells[i, 25], sheet.Cells[i, 27]].Merge();

                                for (int pp2 = 28; pp2 < 35; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int two = 35; two < 36; two++)
                                    sheet.Cells[i, two] = String.Format("2", i, two);

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int pp2 = 40; pp2 < 66; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();

                                for (int dot = 66; dot < 68; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();

                                for (int pdp = 68; pdp < 76; pdp++)
                                    sheet.Cells[i, pdp] = String.Format("ПДП", i, pdp);
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 75]].Merge();

                                for (int pvkr = 76; pvkr < 84; pvkr++)
                                    sheet.Cells[i, pvkr] = String.Format("ПВКР", i, pvkr);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 83]].Merge();

                                for (int zvkr = 84; zvkr < 88; zvkr++)
                                    sheet.Cells[i, zvkr] = String.Format("ЗВКР", i, zvkr);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if(responsGroup.Data[a].Name_Group == "СА50-1-21" 
                                || responsGroup.Data[a].Name_Group == "СА50-2-21" 
                                || responsGroup.Data[a].Name_Group == "СА50-3-21")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();

                                for (int dot = 84; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "СА50-1-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();

                                for (int smia = 82; smia < 86; smia++)
                                    sheet.Cells[i, smia] = String.Format("СМиА", i, smia);
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 85]].Merge();

                                for (int dot = 86; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "СА50-2-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();

                                for (int smia = 76; smia < 80; smia++)
                                    sheet.Cells[i, smia] = String.Format("СМиА", i, smia);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int dot = 86; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "СА50-3-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();

                                for (int smia = 76; smia < 80; smia++)
                                    sheet.Cells[i, smia] = String.Format("СМиА", i, smia);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int dot = 86; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 90; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();

                                for (int equally1 = 90; equally1 < 108; equally1++)
                                    sheet.Cells[i, equally1] = String.Format("=", i, equally1);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "СА50-1-19, СА50-11/1-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();

                                for (int pimls = 26; pimls < 30; pimls++)
                                    sheet.Cells[i, pimls] = String.Format("ПиМЛС", i, pimls);
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int acoc = 40; acoc < 44; acoc++)
                                    sheet.Cells[i, acoc] = String.Format("АСОС", i, acoc);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 43]].Merge();

                                for (int oaocl = 44; oaocl < 47; oaocl++)
                                    sheet.Cells[i, oaocl] = String.Format("ОАОСL", i, oaocl);
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 46]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();

                                for (int pp2 = 69; pp2 < 72; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();

                                for (int pimls = 72; pimls < 80; pimls++)
                                    sheet.Cells[i, pimls] = String.Format("ДиОСВТ", i, pimls);
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 79]].Merge();

                                for (int pp2 = 80; pp2 < 88; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int dot = 88; dot < 90; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();

                                for (int equally1 = 90; equally1 < 108; equally1++)
                                    sheet.Cells[i, equally1] = String.Format("=", i, equally1);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "СА50-2-19")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();

                                for (int pimls = 30; pimls < 34; pimls++)
                                    sheet.Cells[i, pimls] = String.Format("ПиМЛС", i, pimls);
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int oaocl = 40; oaocl < 43; oaocl++)
                                    sheet.Cells[i, oaocl] = String.Format("ОАОСL", i, oaocl);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 42]].Merge();

                                for (int acoc = 43; acoc < 47; acoc++)
                                    sheet.Cells[i, acoc] = String.Format("АСОС", i, acoc);
                                sheet.Range[sheet.Cells[i, 43], sheet.Cells[i, 46]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();

                                for (int pp2 = 69; pp2 < 80; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();

                                for (int pimls = 80; pimls < 88; pimls++)
                                    sheet.Cells[i, pimls] = String.Format("ДиОСВТ", i, pimls);
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 87]].Merge();

                                for (int dot = 88; dot < 90; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();

                                for (int equally1 = 90; equally1 < 108; equally1++)
                                    sheet.Cells[i, equally1] = String.Format("=", i, equally1);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "СА50-3-19")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();

                                for (int pimls = 30; pimls < 34; pimls++)
                                    sheet.Cells[i, pimls] = String.Format("ПиМЛС", i, pimls);
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();

                                for (int acoc = 48; acoc < 52; acoc++)
                                    sheet.Cells[i, acoc] = String.Format("АСОС", i, acoc);
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 51]].Merge();

                                for (int oaocl = 52; oaocl < 55; oaocl++)
                                    sheet.Cells[i, oaocl] = String.Format("ОАОСL", i, oaocl);
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 54]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();

                                for (int pimls = 64; pimls < 68; pimls++)
                                    sheet.Cells[i, pimls] = String.Format("ДиОСВТ", i, pimls);
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 67]].Merge();

                                for (int pp2 = 69; pp2 < 88; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int dot = 88; dot < 90; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();

                                for (int equally1 = 90; equally1 < 108; equally1++)
                                    sheet.Cells[i, equally1] = String.Format("=", i, equally1);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "СА50-4-19")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();

                                for (int pimls = 30; pimls < 34; pimls++)
                                    sheet.Cells[i, pimls] = String.Format("ПиМЛС", i, pimls);
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();

                                for (int oaocl = 49; oaocl < 52; oaocl++)
                                    sheet.Cells[i, oaocl] = String.Format("ОАОСL", i, oaocl);
                                sheet.Range[sheet.Cells[i, 49], sheet.Cells[i, 51]].Merge();

                                for (int acoc = 52; acoc < 56; acoc++)
                                    sheet.Cells[i, acoc] = String.Format("АСОС", i, acoc);
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();

                                for (int pimls = 64; pimls < 68; pimls++)
                                    sheet.Cells[i, pimls] = String.Format("ДиОСВТ", i, pimls);
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 67]].Merge();

                                for (int pp2 = 69; pp2 < 88; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int dot = 88; dot < 90; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();

                                for (int equally1 = 90; equally1 < 108; equally1++)
                                    sheet.Cells[i, equally1] = String.Format("=", i, equally1);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "СА50-5-19, СА50-11/5-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();

                                for (int pimls = 30; pimls < 34; pimls++)
                                    sheet.Cells[i, pimls] = String.Format("ПиМЛС", i, pimls);
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();

                                for (int acoc = 59; acoc < 63; acoc++)
                                    sheet.Cells[i, acoc] = String.Format("АСОС", i, acoc);
                                sheet.Range[sheet.Cells[i, 59], sheet.Cells[i, 62]].Merge();

                                for (int oaocl = 63; oaocl < 65; oaocl++)
                                    sheet.Cells[i, oaocl] = String.Format("ОАОСL", i, oaocl);
                                sheet.Range[sheet.Cells[i, 63], sheet.Cells[i, 64]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();

                                for (int pp2 = 69; pp2 < 80; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();

                                for (int pimls = 80; pimls < 88; pimls++)
                                    sheet.Cells[i, pimls] = String.Format("ДиОСВТ", i, pimls);
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 87]].Merge();

                                for (int dot = 88; dot < 90; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();

                                for (int equally1 = 90; equally1 < 108; equally1++)
                                    sheet.Cells[i, equally1] = String.Format("=", i, equally1);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "СА50-1-18")
                            {
                                for (int osi = 2; osi < 7; osi++)
                                    sheet.Cells[i, osi] = String.Format("ЭОСИ", i, osi);
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 6]].Merge();

                                for (int pp2 = 7; pp2 < 31; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int free = 35; free < 36; free++)
                                    sheet.Cells[i, free] = String.Format("3", i, free);

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int bis = 40; bis < 43; bis++)
                                    sheet.Cells[i, bis] = String.Format("БИС", i, bis);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 42]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();

                                for (int pp2 = 50; pp2 < 67; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();

                                for (int free = 67; free < 68; free++)
                                    sheet.Cells[i, free] = String.Format("3", i, free);

                                for (int pdp = 68; pdp < 76; pdp++)
                                    sheet.Cells[i, pdp] = String.Format("ПДП", i, pdp);
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 75]].Merge();

                                for (int pvkr = 76; pvkr < 84; pvkr++)
                                    sheet.Cells[i, pvkr] = String.Format("ПВКР", i, pvkr);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 83]].Merge();

                                for (int zvkr = 84; zvkr < 88; zvkr++)
                                    sheet.Cells[i, zvkr] = String.Format("ЗВКР", i, zvkr);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "СА50-2-18")
                            {
                                for (int pp2 = 2; pp2 < 7; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();

                                for (int osi = 7; osi < 12; osi++)
                                    sheet.Cells[i, osi] = String.Format("ЭОСИ", i, osi);
                                sheet.Range[sheet.Cells[i, 7], sheet.Cells[i, 11]].Merge();

                                for (int pp2 = 12; pp2 < 31; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int free = 35; free < 36; free++)
                                    sheet.Cells[i, free] = String.Format("3", i, free);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int bis = 43; bis < 46; bis++)
                                    sheet.Cells[i, bis] = String.Format("БИС", i, bis);
                                sheet.Range[sheet.Cells[i, 43], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();

                                for (int pp2 = 50; pp2 < 67; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();

                                for (int free = 67; free < 68; free++)
                                    sheet.Cells[i, free] = String.Format("3", i, free);

                                for (int pdp = 68; pdp < 76; pdp++)
                                    sheet.Cells[i, pdp] = String.Format("ПДП", i, pdp);
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 75]].Merge();

                                for (int pvkr = 76; pvkr < 84; pvkr++)
                                    sheet.Cells[i, pvkr] = String.Format("ПВКР", i, pvkr);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 83]].Merge();

                                for (int zvkr = 84; zvkr < 88; zvkr++)
                                    sheet.Cells[i, zvkr] = String.Format("ЗВКР", i, zvkr);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "СА50-3-18, СА50-11-19")
                            {
                                for (int pp2 = 2; pp2 < 12; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();

                                for (int osi = 12; osi < 17; osi++)
                                    sheet.Cells[i, osi] = String.Format("ЭОСИ", i, osi);
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 16]].Merge();

                                for (int pp2 = 17; pp2 < 31; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int free = 35; free < 36; free++)
                                    sheet.Cells[i, free] = String.Format("3", i, free);

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();

                                for (int bis = 46; bis < 49; bis++)
                                    sheet.Cells[i, bis] = String.Format("БИС", i, bis);
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 48]].Merge();

                                for (int pp2 = 50; pp2 < 67; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();

                                for (int free = 67; free < 68; free++)
                                    sheet.Cells[i, free] = String.Format("3", i, free);

                                for (int pdp = 68; pdp < 76; pdp++)
                                    sheet.Cells[i, pdp] = String.Format("ПДП", i, pdp);
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 75]].Merge();

                                for (int pvkr = 76; pvkr < 84; pvkr++)
                                    sheet.Cells[i, pvkr] = String.Format("ПВКР", i, pvkr);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 83]].Merge();

                                for (int zvkr = 84; zvkr < 88; zvkr++)
                                    sheet.Cells[i, zvkr] = String.Format("ЗВКР", i, zvkr);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "ИСиП-1-21 - ИСиП-16-21")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();

                                for (int dot = 84; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "ИС50-1-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();

                                for (int yp = 68; yp < 74; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.06.01", i, yp);
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int dot = 86; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "ИС50-2-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();

                                for (int yp = 74; yp < 80; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.06.01", i, yp);
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int dot = 86; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "ИС50-3-20, ИС50-11-21")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();

                                for (int yp = 80; yp < 86; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.06.01", i, yp);
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 85]].Merge();

                                for (int dot = 86; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "ИС50-1-19")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();

                                for (int yp = 16; yp < 22; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.07.01", i, yp);
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();

                                for (int yp = 64; yp < 66; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.03.01", i, yp);
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();

                                for (int pp2 = 70; pp2 < 88; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int dot = 88; dot < 90; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();

                                for (int equally1 = 90; equally1 < 108; equally1++)
                                    sheet.Cells[i, equally1] = String.Format("=", i, equally1);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "ИС50-2-19")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();

                                for (int yp = 22; yp < 28; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.07.01", i, yp);
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();

                                for (int yp = 66; yp < 68; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.03.01", i, yp);
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();

                                for (int pp2 = 70; pp2 < 88; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int dot = 88; dot < 90; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();

                                for (int equally1 = 90; equally1 < 108; equally1++)
                                    sheet.Cells[i, equally1] = String.Format("=", i, equally1);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "ИС50-3-19, ИС50-11/3-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();

                                for (int yp = 28; yp < 34; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.07.01", i, yp);
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();

                                for (int yp = 68; yp < 70; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.03.01", i, yp);
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();

                                for (int pp2 = 70; pp2 < 88; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int dot = 88; dot < 108; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();

                                for (int equally = 90; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "ИС50-1-18")
                            {
                                for (int yp = 2; yp < 6; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.05.01", i, yp);
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 5]].Merge();

                                for (int yp = 6; yp < 10; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.02.01", i, yp);
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();

                                for (int pp2 = 14; pp2 < 34; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int pp = 40; pp < 42; pp++)
                                    sheet.Cells[i, pp] = String.Format("ПП.07.01", i, pp);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();

                                for (int pp = 42; pp < 46; pp++)
                                    sheet.Cells[i, pp] = String.Format("ПП.02.01", i, pp);
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();

                                for (int pp = 52; pp < 58; pp++)
                                    sheet.Cells[i, pp] = String.Format("ПП.05.01", i, pp);
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();

                                for (int dot = 66; dot < 68; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();

                                for (int pdp = 68; pdp < 76; pdp++)
                                    sheet.Cells[i, pdp] = String.Format("ПДП", i, pdp);
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 75]].Merge();

                                for (int pvkr = 76; pvkr < 84; pvkr++)
                                    sheet.Cells[i, pvkr] = String.Format("ПВКР", i, pvkr);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 83]].Merge();

                                for (int zvkr = 84; zvkr < 88; zvkr++)
                                    sheet.Cells[i, zvkr] = String.Format("ЗВКР", i, zvkr);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "ИС50-2-18")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();

                                for (int yp = 4; yp < 8; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.05.01", i, yp);
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 7]].Merge();

                                for (int yp = 8; yp < 12; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.02.01", i, yp);
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();

                                for (int pp2 = 14; pp2 < 34; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();

                                for (int pp = 46; pp < 48; pp++)
                                    sheet.Cells[i, pp] = String.Format("ПП.07.01", i, pp);
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();

                                for (int pp = 48; pp < 52; pp++)
                                    sheet.Cells[i, pp] = String.Format("ПП.02.01", i, pp);
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();

                                for (int pp = 58; pp < 64; pp++)
                                    sheet.Cells[i, pp] = String.Format("ПП.05.01", i, pp);
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();

                                for (int dot = 66; dot < 68; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();

                                for (int pdp = 68; pdp < 76; pdp++)
                                    sheet.Cells[i, pdp] = String.Format("ПДП", i, pdp);
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 75]].Merge();

                                for (int pvkr = 76; pvkr < 84; pvkr++)
                                    sheet.Cells[i, pvkr] = String.Format("ПВКР", i, pvkr);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 83]].Merge();

                                for (int zvkr = 84; zvkr < 88; zvkr++)
                                    sheet.Cells[i, zvkr] = String.Format("ЗВКР", i, zvkr);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "ИС50-11-19")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();

                                for (int yp = 6; yp < 10; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.05.01", i, yp);
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 9]].Merge();

                                for (int yp = 10; yp < 14; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.02.01", i, yp);
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 13]].Merge();

                                for (int pp2 = 14; pp2 < 34; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();

                                for (int pp = 46; pp < 48; pp++)
                                    sheet.Cells[i, pp] = String.Format("ПП.07.01", i, pp);
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();

                                for (int pp = 48; pp < 52; pp++)
                                    sheet.Cells[i, pp] = String.Format("ПП.02.01", i, pp);
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();

                                for (int pp = 58; pp < 64; pp++)
                                    sheet.Cells[i, pp] = String.Format("ПП.05.01", i, pp);
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();

                                for (int dot = 66; dot < 68; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();

                                for (int pdp = 68; pdp < 76; pdp++)
                                    sheet.Cells[i, pdp] = String.Format("ПДП", i, pdp);
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 75]].Merge();

                                for (int pvkr = 76; pvkr < 84; pvkr++)
                                    sheet.Cells[i, pvkr] = String.Format("ПВКР", i, pvkr);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 83]].Merge();

                                for (int zvkr = 84; zvkr < 88; zvkr++)
                                    sheet.Cells[i, zvkr] = String.Format("ЗВКР", i, zvkr);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "П50-1-20" 
                                || responsGroup.Data[a].Name_Group == "П50-2-20"
                                || responsGroup.Data[a].Name_Group == "П50-5-20" 
                                || responsGroup.Data[a].Name_Group == "П50-6-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();

                                for (int prp = 62; prp < 66; prp++)
                                    sheet.Cells[i, prp] = String.Format("ПрП", i, prp);
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int dot = 88; dot < 90; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();

                                for (int equally1 = 90; equally1 < 108; equally1++)
                                    sheet.Cells[i, equally1] = String.Format("=", i, equally1);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "П50-3-20" 
                                || responsGroup.Data[a].Name_Group == "П50-4-20"
                                || responsGroup.Data[a].Name_Group == "П50-7-20" 
                                || responsGroup.Data[a].Name_Group == "П50-11-21" 
                                || responsGroup.Data[a].Name_Group == "Т50-1-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();

                                for (int prp = 58; prp < 62; prp++)
                                    sheet.Cells[i, prp] = String.Format("ПрП", i, prp);
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int dot = 88; dot < 90; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();

                                for (int equally1 = 90; equally1 < 108; equally1++)
                                    sheet.Cells[i, equally1] = String.Format("=", i, equally1);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "П50-1-19, П50-11/1-20" 
                                || responsGroup.Data[a].Name_Group == "П50-2-19")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();

                                for (int rpm = 32; rpm < 34; rpm++)
                                    sheet.Cells[i, rpm] = String.Format("РПМ", i, rpm);
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int riis = 40; riis < 44; riis++)
                                    sheet.Cells[i, riis] = String.Format("РиЭИС", i, riis);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 43]].Merge();

                                for (int pp2 = 44; pp2 < 50; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();

                                for (int trpo = 50; trpo < 56; trpo++)
                                    sheet.Cells[i, trpo] = String.Format("ТРПО", i, trpo);
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 55]].Merge();

                                for (int pp2 = 56; pp2 < 86; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int dot = 86; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "П50-3-19, П50-11/3-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();

                                for (int rpm = 30; rpm < 32; rpm++)
                                    sheet.Cells[i, rpm] = String.Format("РПМ", i, rpm);
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int pp2 = 40; pp2 < 46; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();

                                for (int riis = 46; riis < 50; riis++)
                                    sheet.Cells[i, riis] = String.Format("РиЭИС", i, riis);
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 49]].Merge();

                                for (int trpo = 50; trpo < 56; trpo++)
                                    sheet.Cells[i, trpo] = String.Format("ТРПО", i, trpo);
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 55]].Merge();

                                for (int pp2 = 56; pp2 < 86; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int dot = 86; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "П50-4-19, П50-11/4-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();

                                for (int rpm = 30; rpm < 32; rpm++)
                                    sheet.Cells[i, rpm] = String.Format("РПМ", i, rpm);
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int trpo = 40; trpo < 46; trpo++)
                                    sheet.Cells[i, trpo] = String.Format("ТРПО", i, trpo);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 45]].Merge();

                                for (int riis = 46; riis < 50; riis++)
                                    sheet.Cells[i, riis] = String.Format("РиЭИС", i, riis);
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 49]].Merge();

                                for (int pp2 = 50; pp2 < 86; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int dot = 86; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "П50-5-19, П50-11/5-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();

                                for (int rpm = 28; rpm < 30; rpm++)
                                    sheet.Cells[i, rpm] = String.Format("РПМ", i, rpm);
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int trpo = 40; trpo < 46; trpo++)
                                    sheet.Cells[i, trpo] = String.Format("ТРПО", i, trpo);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 45]].Merge();

                                for (int pp2 = 46; pp2 < 52; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();

                                for (int riis = 52; riis < 56; riis++)
                                    sheet.Cells[i, riis] = String.Format("РиЭИС", i, riis);
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 55]].Merge();

                                for (int pp2 = 56; pp2 < 86; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int dot = 86; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "П50-6-19, П50-11/6-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();

                                for (int rpm = 28; rpm < 30; rpm++)
                                    sheet.Cells[i, rpm] = String.Format("РПМ", i, rpm);
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int trpo = 40; trpo < 46; trpo++)
                                    sheet.Cells[i, trpo] = String.Format("ТРПО", i, trpo);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 45]].Merge();

                                for (int pp2 = 46; pp2 < 48; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();

                                for (int riis = 48; riis < 52; riis++)
                                    sheet.Cells[i, riis] = String.Format("РиЭИС", i, riis);
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 51]].Merge();

                                for (int pp2 = 52; pp2 < 86; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int dot = 86; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "П50-1-18" 
                                || responsGroup.Data[a].Name_Group == "П50-5-18")
                            {
                                for (int pp2 = 6; pp2 < 22; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();

                                for (int vippo = 22; vippo < 28; vippo++)
                                    sheet.Cells[i, vippo] = String.Format("ВиППО", i, vippo);
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 27]].Merge();

                                for (int pp2 = 28; pp2 < 34; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int pp2 = 40; pp2 < 66; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();

                                for (int dot = 66; dot < 68; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();

                                for (int pdp = 68; pdp < 76; pdp++)
                                    sheet.Cells[i, pdp] = String.Format("ПДП", i, pdp);
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 75]].Merge();

                                for (int pvkr = 76; pvkr < 84; pvkr++)
                                    sheet.Cells[i, pvkr] = String.Format("ПВКР", i, pvkr);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 83]].Merge();

                                for (int zvkr = 84; zvkr < 88; zvkr++)
                                    sheet.Cells[i, zvkr] = String.Format("ЗВКР", i, zvkr);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "П50-2-18" 
                                || responsGroup.Data[a].Name_Group == "П50-6-18")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();

                                for (int pp2 = 6; pp2 < 16; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();

                                for (int vippo = 16; vippo < 22; vippo++)
                                    sheet.Cells[i, vippo] = String.Format("ВиППО", i, vippo);
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 21]].Merge();

                                for (int pp2 = 22; pp2 < 34; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int pp2 = 40; pp2 < 66; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();

                                for (int dot = 66; dot < 68; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();

                                for (int pdp = 68; pdp < 76; pdp++)
                                    sheet.Cells[i, pdp] = String.Format("ПДП", i, pdp);
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 75]].Merge();

                                for (int pvkr = 76; pvkr < 84; pvkr++)
                                    sheet.Cells[i, pvkr] = String.Format("ПВКР", i, pvkr);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 83]].Merge();

                                for (int zvkr = 84; zvkr < 88; zvkr++)
                                    sheet.Cells[i, zvkr] = String.Format("ЗВКР", i, zvkr);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "П50-3-18" 
                                || responsGroup.Data[a].Name_Group == "Т50-1-18")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();

                                for (int pp2 = 6; pp2 < 10; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();

                                for (int vippo = 10; vippo < 16; vippo++)
                                    sheet.Cells[i, vippo] = String.Format("ВиППО", i, vippo);
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 15]].Merge();

                                for (int pp2 = 16; pp2 < 34; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int pp2 = 40; pp2 < 66; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();

                                for (int dot = 66; dot < 68; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();

                                for (int pdp = 68; pdp < 76; pdp++)
                                    sheet.Cells[i, pdp] = String.Format("ПДП", i, pdp);
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 75]].Merge();

                                for (int pvkr = 76; pvkr < 84; pvkr++)
                                    sheet.Cells[i, pvkr] = String.Format("ПВКР", i, pvkr);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 83]].Merge();

                                for (int zvkr = 84; zvkr < 88; zvkr++)
                                    sheet.Cells[i, zvkr] = String.Format("ЗВКР", i, zvkr);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "П50-4-18")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();

                                for (int vippo = 4; vippo < 10; vippo++)
                                    sheet.Cells[i, vippo] = String.Format("ВиППО", i, vippo);
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();

                                for (int pp2 = 12; pp2 < 34; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int pp2 = 40; pp2 < 66; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();

                                for (int dot = 66; dot < 68; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();

                                for (int pdp = 68; pdp < 76; pdp++)
                                    sheet.Cells[i, pdp] = String.Format("ПДП", i, pdp);
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 75]].Merge();

                                for (int pvkr = 76; pvkr < 84; pvkr++)
                                    sheet.Cells[i, pvkr] = String.Format("ПВКР", i, pvkr);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 83]].Merge();

                                for (int zvkr = 84; zvkr < 88; zvkr++)
                                    sheet.Cells[i, zvkr] = String.Format("ЗВКР", i, zvkr);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "Т50-1-19, Т50-11-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();

                                for (int rpm = 32; rpm < 34; rpm++)
                                    sheet.Cells[i, rpm] = String.Format("РПМ", i, rpm);
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int trpo = 40; trpo < 46; trpo++)
                                    sheet.Cells[i, trpo] = String.Format("ТРПО", i, trpo);
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 46]].Merge();

                                for (int pp2 = 46; pp2 < 60; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();

                                for (int riis = 60; riis < 64; riis++)
                                    sheet.Cells[i, riis] = String.Format("РиЭИС", i, riis);
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 63]].Merge();

                                for (int pp2 = 64; pp2 < 86; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int dot = 86; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "БД50-1-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();

                                for (int yp = 16; yp < 24; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.02.01", i, yp);
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();

                                for (int yp = 64; yp < 68; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.01.01", i, yp);
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();

                                for (int yp = 82; yp < 86; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.04.01", i, yp);
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 85]].Merge();

                                for (int dot = 86; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "БД50-1-19, БД50-11/1-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();

                                for (int yp = 58; yp < 64; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.07.01", i, yp);
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 63]].Merge();

                                for (int yp = 64; yp < 70; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.11.01", i, yp);
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int dot = 88; dot < 90; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();

                                for (int equally1 = 90; equally1 < 108; equally1++)
                                    sheet.Cells[i, equally1] = String.Format("=", i, equally1);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "БД50-1-18")
                            {
                                for (int pp2 = 2; pp2 < 34; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();

                                for (int pp = 52; pp < 55; pp++)
                                    sheet.Cells[i, pp] = String.Format("ПП.07.01", i, pp);
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 54]].Merge();

                                for (int pp = 55; pp < 58; pp++)
                                    sheet.Cells[i, pp] = String.Format("ПП.11.01", i, pp);
                                sheet.Range[sheet.Cells[i, 55], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();

                                for (int dot = 66; dot < 68; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();

                                for (int pdp = 68; pdp < 76; pdp++)
                                    sheet.Cells[i, pdp] = String.Format("ПДП", i, pdp);
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 75]].Merge();

                                for (int pvkr = 76; pvkr < 84; pvkr++)
                                    sheet.Cells[i, pvkr] = String.Format("ПВКР", i, pvkr);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 83]].Merge();

                                for (int zvkr = 84; zvkr < 88; zvkr++)
                                    sheet.Cells[i, zvkr] = String.Format("ЗВКР", i, zvkr);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "ВД50-1-20, ВД50-11/1-21")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();

                                for (int yp = 26; yp < 28; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.05.01", i, yp);
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();

                                for (int dot = 28; dot < 30; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();

                                for (int yp = 30; yp < 36; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.05.01", i, yp);
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();

                                for (int yp = 56; yp < 64; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.08.01", i, yp);
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int dot = 86; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "ВД50-2-20, ВД50-11/2-21")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();

                                for (int dot = 26; dot < 28; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();

                                for (int yp = 28; yp < 36; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.05.01", i, yp);
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();

                                for (int yp = 56; yp < 64; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.08.01", i, yp);
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int dot = 86; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "ВД50-3-20, ВД50-11/3-21")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();

                                for (int yp = 26; yp < 30; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.05.01", i, yp);
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 29]].Merge();

                                for (int dot = 30; dot < 32; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();

                                for (int yp = 32; yp < 36; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.05.01", i, yp);
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();

                                for (int yp = 56; yp < 64; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.08.01", i, yp);
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int dot = 86; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "ВД50-4-20, ВД50-11/4-21")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();

                                for (int yp = 26; yp < 34; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.05.01", i, yp);
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();

                                for (int yp = 56; yp < 64; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.08.01", i, yp);
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int dot = 86; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "ВД50-1-19, ВД50-11/1-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();

                                for (int dot = 76; dot < 78; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();

                                for (int yp = 78; yp < 90; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.09.01", i, yp);
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 89]].Merge();

                                for (int equally = 90; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "ВД50-2-19, ВД50-11/2-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();

                                for (int yp = 76; yp < 88; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.09.01", i, yp);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 87]].Merge();

                                for (int dot = 88; dot < 90; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();

                                for (int equally = 90; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "ВД50-3-19, ВД50-11/3-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();

                                for (int yp = 76; yp < 84; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.09.01", i, yp);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 83]].Merge();

                                for (int dot = 84; dot < 86; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int yp = 86; yp < 90; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.09.01", i, yp);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 89]].Merge();

                                for (int equally = 90; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "ВД50-1-18" 
                                || responsGroup.Data[a].Name_Group == "ВД50-2-18, ВД50-11-19" 
                                || responsGroup.Data[a].Name_Group == "ВД50-3-18" 
                                || responsGroup.Data[a].Name_Group == "ВД50-4-18" 
                                || responsGroup.Data[a].Name_Group == "ВД50-5-18" 
                                || responsGroup.Data[a].Name_Group == "БИ50-1-18" 
                                || responsGroup.Data[a].Name_Group == "БИ50-2-18, БИ50-11-19")
                            {
                                for (int pp2 = 2; pp2 < 34; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();

                                for (int pp2 = 54; pp2 < 66; pp2++)
                                    sheet.Cells[i, pp2] = String.Format(" ", i, pp2);
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();

                                for (int dot = 66; dot < 68; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();

                                for (int pdp = 68; pdp < 76; pdp++)
                                    sheet.Cells[i, pdp] = String.Format("ПДП", i, pdp);
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 75]].Merge();

                                for (int pvkr = 76; pvkr < 84; pvkr++)
                                    sheet.Cells[i, pvkr] = String.Format("ПВКР", i, pvkr);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 83]].Merge();

                                for (int zvkr = 84; zvkr < 88; zvkr++)
                                    sheet.Cells[i, zvkr] = String.Format("ЗВКР", i, zvkr);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "Ю-1-21" 
                                || responsGroup.Data[a].Name_Group == "Ю-2-21")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();

                                for (int dot = 84; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "Ю-1-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();

                                for (int yp = 60; yp < 66; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.01.01", i, yp);
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();

                                for (int dot = 82; dot < 84; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int pp = 84; pp < 88; pp++)
                                    sheet.Cells[i, pp] = String.Format("ПП.02.01", i, pp);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "Ю-11-21")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();

                                for (int yp = 62; yp < 68; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.01.01", i, yp);
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();

                                for (int dot = 82; dot < 84; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();

                                for (int pp = 84; pp < 88; pp++)
                                    sheet.Cells[i, pp] = String.Format("ПП.02.01", i, pp);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "Ю-1-19" 
                                || responsGroup.Data[a].Name_Group == "Ю-11-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();
                                
                                for (int free = 34; free < 35; free++)
                                    sheet.Cells[i, free] = String.Format("3", i, free);

                                for (int yp = 35; yp < 36; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.02.01", i, yp);

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();

                                for (int yp = 40; yp < 41; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.02.01", i, yp);
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                
                                for (int pp = 66; pp < 67; pp++)
                                    sheet.Cells[i, pp] = String.Format("ПП.02.02", i, pp);

                                for (int one = 67; one < 68; one++)
                                    sheet.Cells[i, one] = String.Format("1", i, one);

                                for (int pdp = 68; pdp < 76; pdp++)
                                    sheet.Cells[i, pdp] = String.Format("ПДП", i, pdp);
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 75]].Merge();

                                for (int pvkr = 76; pvkr < 84; pvkr++)
                                    sheet.Cells[i, pvkr] = String.Format("ПВКР", i, pvkr);
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 83]].Merge();

                                for (int zvkr = 84; zvkr < 88; zvkr++)
                                    sheet.Cells[i, zvkr] = String.Format("ЗВКР", i, zvkr);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "БИ50-1-21" 
                                || responsGroup.Data[a].Name_Group == "БИ50-2-21"
                                || responsGroup.Data[a].Name_Group == "БИ50-3-21")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();

                                for (int dot = 84; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "БИ50-1-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();

                                for (int yp = 6; yp < 14; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.04.01", i, yp);
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();

                                for (int dot = 78; dot < 80; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();

                                for (int yp = 80; yp < 88; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.01.01", i, yp);
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "БИ50-2-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();

                                for (int yp = 6; yp < 14; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.04.01", i, yp);
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();

                                for (int yp = 78; yp < 86; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.01.01", i, yp);
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 85]].Merge();

                                for (int dot = 86; dot < 88; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "БИ50-3-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();

                                for (int yp = 10; yp < 18; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.04.01", i, yp);
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();

                                for (int yp = 78; yp < 82; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.01.01", i, yp);
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 81]].Merge();

                                for (int dot = 82; dot < 84; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();

                                for (int yp = 84; yp < 88; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.01.01", i, yp);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "БИ50-11-21")
                            {
                                for (int yp = 2; yp < 10; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.04.01", i, yp);
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 65]].Merge();
                                sheet.Range[sheet.Cells[i, 66], sheet.Cells[i, 67]].Merge();
                                sheet.Range[sheet.Cells[i, 68], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();

                                for (int yp = 78; yp < 82; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.01.01", i, yp);
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 81]].Merge();

                                for (int dot = 82; dot < 84; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();

                                for (int yp = 84; yp < 88; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.01.01", i, yp);
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 87]].Merge();

                                for (int equally = 88; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "БИ50-1-19, БИ50-11/1-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 59]].Merge();
                                sheet.Range[sheet.Cells[i, 60], sheet.Cells[i, 61]].Merge();
                                sheet.Range[sheet.Cells[i, 62], sheet.Cells[i, 63]].Merge();

                                for (int yp = 64; yp < 70; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.03.01", i, yp);
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();

                                for (int yp = 82; yp < 88; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.02.01", i, yp);
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 87]].Merge();

                                for (int dot = 88; dot < 90; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();

                                for (int equally = 90; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "БИ50-2-19, БИ50-11/2-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();

                                for (int yp = 58; yp < 64; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.03.01", i, yp);
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 63]].Merge();

                                for (int yp = 64; yp < 70; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.02.01", i, yp);
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int dot = 88; dot < 90; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();

                                for (int equally = 90; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            else if (responsGroup.Data[a].Name_Group == "БИ50-3-19, БИ50-11/3-20")
                            {
                                sheet.Range[sheet.Cells[i, 2], sheet.Cells[i, 3]].Merge();
                                sheet.Range[sheet.Cells[i, 4], sheet.Cells[i, 5]].Merge();
                                sheet.Range[sheet.Cells[i, 6], sheet.Cells[i, 7]].Merge();
                                sheet.Range[sheet.Cells[i, 8], sheet.Cells[i, 9]].Merge();
                                sheet.Range[sheet.Cells[i, 10], sheet.Cells[i, 11]].Merge();
                                sheet.Range[sheet.Cells[i, 12], sheet.Cells[i, 13]].Merge();
                                sheet.Range[sheet.Cells[i, 14], sheet.Cells[i, 15]].Merge();
                                sheet.Range[sheet.Cells[i, 16], sheet.Cells[i, 17]].Merge();
                                sheet.Range[sheet.Cells[i, 18], sheet.Cells[i, 19]].Merge();
                                sheet.Range[sheet.Cells[i, 20], sheet.Cells[i, 21]].Merge();
                                sheet.Range[sheet.Cells[i, 22], sheet.Cells[i, 23]].Merge();
                                sheet.Range[sheet.Cells[i, 24], sheet.Cells[i, 25]].Merge();
                                sheet.Range[sheet.Cells[i, 26], sheet.Cells[i, 27]].Merge();
                                sheet.Range[sheet.Cells[i, 28], sheet.Cells[i, 29]].Merge();
                                sheet.Range[sheet.Cells[i, 30], sheet.Cells[i, 31]].Merge();
                                sheet.Range[sheet.Cells[i, 32], sheet.Cells[i, 33]].Merge();

                                for (int dot = 34; dot < 36; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 34], sheet.Cells[i, 35]].Merge();

                                for (int equally = 36; equally < 40; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 36], sheet.Cells[i, 37]].Merge();
                                sheet.Range[sheet.Cells[i, 38], sheet.Cells[i, 39]].Merge();
                                sheet.Range[sheet.Cells[i, 40], sheet.Cells[i, 41]].Merge();
                                sheet.Range[sheet.Cells[i, 42], sheet.Cells[i, 43]].Merge();
                                sheet.Range[sheet.Cells[i, 44], sheet.Cells[i, 45]].Merge();
                                sheet.Range[sheet.Cells[i, 46], sheet.Cells[i, 47]].Merge();
                                sheet.Range[sheet.Cells[i, 48], sheet.Cells[i, 49]].Merge();
                                sheet.Range[sheet.Cells[i, 50], sheet.Cells[i, 51]].Merge();
                                sheet.Range[sheet.Cells[i, 52], sheet.Cells[i, 53]].Merge();
                                sheet.Range[sheet.Cells[i, 54], sheet.Cells[i, 55]].Merge();
                                sheet.Range[sheet.Cells[i, 56], sheet.Cells[i, 57]].Merge();

                                for (int yp = 58; yp < 64; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.02.01", i, yp);
                                sheet.Range[sheet.Cells[i, 58], sheet.Cells[i, 63]].Merge();

                                for (int yp = 64; yp < 70; yp++)
                                    sheet.Cells[i, yp] = String.Format("УП.03.01", i, yp);
                                sheet.Range[sheet.Cells[i, 64], sheet.Cells[i, 69]].Merge();
                                sheet.Range[sheet.Cells[i, 70], sheet.Cells[i, 71]].Merge();
                                sheet.Range[sheet.Cells[i, 72], sheet.Cells[i, 73]].Merge();
                                sheet.Range[sheet.Cells[i, 74], sheet.Cells[i, 75]].Merge();
                                sheet.Range[sheet.Cells[i, 76], sheet.Cells[i, 77]].Merge();
                                sheet.Range[sheet.Cells[i, 78], sheet.Cells[i, 79]].Merge();
                                sheet.Range[sheet.Cells[i, 80], sheet.Cells[i, 81]].Merge();
                                sheet.Range[sheet.Cells[i, 82], sheet.Cells[i, 83]].Merge();
                                sheet.Range[sheet.Cells[i, 84], sheet.Cells[i, 85]].Merge();
                                sheet.Range[sheet.Cells[i, 86], sheet.Cells[i, 87]].Merge();

                                for (int dot = 88; dot < 90; dot++)
                                    sheet.Cells[i, dot] = String.Format("::", i, dot);
                                sheet.Range[sheet.Cells[i, 88], sheet.Cells[i, 89]].Merge();

                                for (int equally = 90; equally < 108; equally++)
                                    sheet.Cells[i, equally] = String.Format("=", i, equally);
                                sheet.Range[sheet.Cells[i, 90], sheet.Cells[i, 91]].Merge();
                                sheet.Range[sheet.Cells[i, 92], sheet.Cells[i, 93]].Merge();
                                sheet.Range[sheet.Cells[i, 94], sheet.Cells[i, 95]].Merge();
                                sheet.Range[sheet.Cells[i, 96], sheet.Cells[i, 97]].Merge();
                                sheet.Range[sheet.Cells[i, 98], sheet.Cells[i, 99]].Merge();
                                sheet.Range[sheet.Cells[i, 100], sheet.Cells[i, 101]].Merge();
                                sheet.Range[sheet.Cells[i, 102], sheet.Cells[i, 103]].Merge();
                                sheet.Range[sheet.Cells[i, 104], sheet.Cells[i, 105]].Merge();
                                sheet.Range[sheet.Cells[i, 106], sheet.Cells[i, 107]].Merge();
                            }
                            a++;
                        }
                    }
                }
                catch
                {

                }
            }
        }
    }
}
