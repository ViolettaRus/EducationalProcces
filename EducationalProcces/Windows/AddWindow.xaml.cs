using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace EducationalProcces.Windows
{
    /// <summary>
    /// Логика взаимодействия для AddWindow.xaml
    /// </summary>
    public partial class AddWindow : Window
    {
        public AddWindow()
        {
            InitializeComponent();
            APIHelper.ActivateLink();           
        }

        private void ExitMainBtn_Click(object sender, RoutedEventArgs e)
        {
            MainWindow mainWindow = new();
            mainWindow.Show();
            this.Close();
        }

        private void AddGroupBtn_Click(object sender, RoutedEventArgs e)
        {
            Group.Visibility = Visibility.Visible;
            Main.Visibility = Visibility.Hidden;
            LoadDataGroup();
        }

        public async void LoadDataGroup()
        {
            ResponseModel<List<Group>> responseGroup = await APIHelper.GetDataAsync<List<Group>>("\"\"", "\"\"", "\"\"", typeof(Group));
            if (responseGroup.StatusCode == 201)
            {
                DataGroup.ItemsSource = responseGroup.Data;
            }
        }

        private async void AddDataGroupBtn_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(NameGroupBox.Text))
            {
                MessageBox.Show("Заполните поля!", " ",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                Group group = new Group();
                ResponseModel<List<Group>> responsGroup = await APIHelper.GetDataAsync<List<Group>>(nameof(group.Name_Group), NameGroupBox.Text, "\"\"", typeof(Group));

                if (responsGroup.StatusCode == 201)
                {
                    if (responsGroup.Data[0].Name_Group == NameGroupBox.Text)
                    {
                        MessageBox.Show("Такая группа уже существует");
                    }
                }
                else
                {
                    ResponseModel<Group> responseGroup = await APIHelper.PostDataAsync<Group>(new Group()
                    {
                        Name_Group = NameGroupBox.Text
                    });
                    LoadDataGroup();
                }
            }
        }

        private async void EditGroupBtn_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(NameGroupBox.Text))
            {
                MessageBox.Show("Заполните поля!", " ",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                Group group = new Group();
                ResponseModel<List<Group>> responsGroup = await APIHelper.GetDataAsync<List<Group>>(nameof(group.Name_Group), NameGroupBox.Text, "\"\"", typeof(Group));

                if (responsGroup.StatusCode == 201)
                {
                    if (responsGroup.Data[0].Name_Group == NameGroupBox.Text)
                    {
                        MessageBox.Show("Такая группа уже существует");
                    }
                }
                else
                {
                    int idGroup;
                    Group groups = (sender as Button).DataContext as Group;
                    idGroup = groups.ID_Group;
                    ResponseModel<Group> responseGroup = await APIHelper.PutDataAsync<Group>(new Group()
                    {
                        ID_Group = idGroup,
                        Name_Group = NameGroupBox.Text

                    });
                    LoadDataGroup();
                }
            }
        }

        private async void DeleteGroupBtn_Click(object sender, RoutedEventArgs e)
        {
            int idGroup;
            Group group = (sender as Button).DataContext as Group;
            idGroup = group.ID_Group;
            ResponseModel<Group> responseGroup = await APIHelper.DeleteDataAsync<Group>(new Group()
            {
                ID_Group = idGroup
            });
            LoadDataGroup();
        }

        private void ExitGroupBtn_Click(object sender, RoutedEventArgs e)
        {
            Group.Visibility = Visibility.Hidden;
            Main.Visibility = Visibility.Visible;
            NameGroupBox.Text = " ";
        }
        //Предмет
        private void AddSubjectBtn_Click(object sender, RoutedEventArgs e)
        {
            Subject.Visibility = Visibility.Visible;
            Main.Visibility = Visibility.Hidden;
            LoadDataSubject();
        }

        public async void LoadDataSubject()
        {
            ResponseModel<List<Subject>> responseSubject = await APIHelper.GetDataAsync<List<Subject>>("\"\"", "\"\"", "\"\"", typeof(Subject));
            if (responseSubject.StatusCode == 201)
            {
                DataSubjcet.ItemsSource = responseSubject.Data;
            }
        }

        private async void AddDataSubjcetBtn_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(NameSubjcetBox.Text))
            {
                MessageBox.Show("Заполните поля!", " ",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                Subject subject = new Subject();
                ResponseModel<List<Subject>> responsSubject = await APIHelper.GetDataAsync<List<Subject>>(nameof(subject.Name_Subject), NameSubjcetBox.Text, "\"\"", typeof(Subject));

                if (responsSubject.StatusCode == 201)
                {
                    if (responsSubject.Data[0].Name_Subject == NameSubjcetBox.Text)
                    {
                        MessageBox.Show("Такой предмет уже существует");
                    }
                }
                else
                {
                    ResponseModel<Subject> responseSubject = await APIHelper.PostDataAsync<Subject>(new Subject()
                    {
                        Name_Subject = NameSubjcetBox.Text
                    });
                    LoadDataSubject();
                }
            }
        }

        private async void DeleteSubjcetBtn_Click(object sender, RoutedEventArgs e)
        {
            int idSubject;
            Subject subject = (sender as Button).DataContext as Subject;
            idSubject = subject.ID_Subject;
            ResponseModel<Subject> responseSubject = await APIHelper.DeleteDataAsync<Subject>(new Subject()
            {
                ID_Subject = idSubject
            });
            LoadDataSubject();
        }

        private async void EditSubjcetBtn_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(NameSubjcetBox.Text))
            {
                MessageBox.Show("Заполните поля!", " ",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                Subject subject = new Subject();
                ResponseModel<List<Subject>> responsSubject = await APIHelper.GetDataAsync<List<Subject>>(nameof(subject.Name_Subject), NameSubjcetBox.Text, "\"\"", typeof(Subject));

                if (responsSubject.StatusCode == 201)
                {
                    if (responsSubject.Data[0].Name_Subject == NameSubjcetBox.Text)
                    {
                        MessageBox.Show("Такой предмет уже существует");
                    }
                }
                else
                {
                    int idSubject;
                    Subject subject1 = (sender as Button).DataContext as Subject;
                    idSubject = subject1.ID_Subject;
                    ResponseModel<Subject> responseSubject = await APIHelper.PutDataAsync<Subject>(new Subject()
                    {
                        ID_Subject = idSubject,
                        Name_Subject = NameSubjcetBox.Text

                    });
                    LoadDataSubject();
                }
            }
        }

        private void ExitSubjcetBtn_Click(object sender, RoutedEventArgs e)
        {
            Subject.Visibility = Visibility.Hidden;
            Main.Visibility = Visibility.Visible;
            NameSubjcetBox.Text = " ";
        }
        //Преподаватель
        private void AddTeachertBtn_Click(object sender, RoutedEventArgs e)
        {
            Teacher.Visibility = Visibility.Visible;
            Main.Visibility = Visibility.Hidden;
            LoadDataTeacher();
        }

        public async void LoadDataTeacher()
        {
            ResponseModel<List<Teacher>> responseTeacher = await APIHelper.GetDataAsync<List<Teacher>>("\"\"", "\"\"", "\"\"", typeof(Teacher));
            if (responseTeacher.StatusCode == 201)
            {
                DataTeacher.ItemsSource = responseTeacher.Data;
            }
        }

        private async void AddDataTeacherBtn_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(FirstNameTeacherBox.Text + NameTeacherBox.Text + MidlleNameTeacherBox.Text))
            {
                MessageBox.Show("Заполните поля!", " ",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                Teacher teacher = new Teacher();
                ResponseModel<List<Teacher>> responsTeacher = await APIHelper.GetDataAsync<List<Teacher>>(nameof(teacher.FIO), FirstNameTeacherBox.Text + " " + NameTeacherBox.Text + " " + MidlleNameTeacherBox.Text, "\"\"", typeof(Teacher));

                if (responsTeacher.StatusCode == 201)
                {
                    if (responsTeacher.Data[0].FIO == FirstNameTeacherBox.Text + " " + NameTeacherBox.Text + " " + MidlleNameTeacherBox.Text)
                    {
                        MessageBox.Show("Такой преподаватель уже существует");
                    }
                }
                else
                {
                    ResponseModel<Teacher> responseTeacher = await APIHelper.PostDataAsync<Teacher>(new Teacher()
                    {
                        FIO = FirstNameTeacherBox.Text + " " + NameTeacherBox.Text + " " + MidlleNameTeacherBox.Text,
                        Phone = PhoneTeacherBox.Text
                    });
                    LoadDataTeacher();
                }
            }
        }

        private async void EditTeacherBtn_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(FirstNameTeacherBox.Text + NameTeacherBox.Text + MidlleNameTeacherBox.Text))
            {
                MessageBox.Show("Заполните поля!", " ",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                Teacher teacher = new Teacher();
                ResponseModel<List<Teacher>> responsTeacher = await APIHelper.GetDataAsync<List<Teacher>>(nameof(teacher.FIO), FirstNameTeacherBox.Text + " " + NameTeacherBox.Text + " " + MidlleNameTeacherBox.Text, "\"\"", typeof(Teacher));

                if (responsTeacher.StatusCode == 201)
                {
                    if (responsTeacher.Data[0].FIO == FirstNameTeacherBox.Text + " " + NameTeacherBox.Text + " " + MidlleNameTeacherBox.Text)
                    {
                        MessageBox.Show("Такой преподаватель уже существует");
                    }
                }
                else
                {
                    int idTeacher;
                    Teacher teacher1 = (sender as Button).DataContext as Teacher;
                    idTeacher = teacher1.ID_Teacher;
                    ResponseModel<Teacher> responseTeacher = await APIHelper.PutDataAsync<Teacher>(new Teacher()
                    {
                        ID_Teacher = idTeacher,
                        FIO = FirstNameTeacherBox.Text + " " + NameTeacherBox.Text + " " + MidlleNameTeacherBox.Text,
                        Phone = PhoneTeacherBox.Text

                    });
                    LoadDataTeacher();
                }
            }
        }

        private async void DeleteTeacherBtn_Click(object sender, RoutedEventArgs e)
        {
            int idTeacher;
            Teacher teacher = (sender as Button).DataContext as Teacher;
            idTeacher = teacher.ID_Teacher;
            ResponseModel<Teacher> responseTeacher = await APIHelper.DeleteDataAsync<Teacher>(new Teacher()
            {
                ID_Teacher = idTeacher
            });
            LoadDataTeacher();
        }

        private void ExitTeacherBtn_Click(object sender, RoutedEventArgs e)
        {
            Teacher.Visibility = Visibility.Hidden;
            Main.Visibility = Visibility.Visible;
            NameTeacherBox.Text = " ";
            MidlleNameTeacherBox.Text = " ";
            FirstNameTeacherBox.Text = " ";
            PhoneTeacherBox.Text = " ";
        }

        private void ExcelBtn_Click(object sender, RoutedEventArgs e)
        {
            //Объявляем приложение
            Excel.Application ex = new Excel.Application();
            //Отобразить Excel
            ex.Visible = true;
            //Количество листов в рабочей книге
            ex.SheetsInNewWorkbook = 2;
            //Добавить рабочую книгу
            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);
            //Отключить отображение окон с сообщениями
            ex.DisplayAlerts = false;
            //Получаем первый лист документа (счет начинается с 1)
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
            //Название листа (вкладки снизу)
            sheet.Name = "Расписание";
            //Пример заполнения ячеек
            for (int i = 1; i <= 9; i++)
            {
                for (int j = 1; j < 9; j++)
                    sheet.Cells[i, j] = String.Format("Boom {0} {1}", i, j);
            }
            ////Захватываем диапазон ячеек
            //Excel.Range range1 = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[9, 9]);
            ////Шрифт для диапазона
            //range1.Cells.Font.Name = "Tahoma";
            ////Размер шрифта для диапазона
            //range1.Cells.Font.Size = 10;
            ////Захватываем другой диапазон ячеек
            //Excel.Range range2 = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[9, 2]);
            //range2.Cells.Font.Name = "Times New Roman";
            
        }
    }
}
