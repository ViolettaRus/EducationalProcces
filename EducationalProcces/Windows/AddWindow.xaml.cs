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
        /// <summary>
        /// Обращение к методу "ActivateLink" из класса "APIHelper"
        /// </summary>
        public AddWindow()
        {
            InitializeComponent();
            APIHelper.ActivateLink();           
        }
        /// <summary>
        /// Закрытие окна "Главная"
        /// Открытие окна "Авторизации"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExitMainBtn_Click(object sender, RoutedEventArgs e)
        {
            MainWindow mainWindow = new();
            mainWindow.Show();
            this.Close();
        }
        ///Группы
        /// <summary>
        /// Открытие панели "Group" 
        /// Закрытие панели "Main"
        /// Обновление данных о группах в DataGroup через API
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddGroupBtn_Click(object sender, RoutedEventArgs e)
        {
            Group.Visibility = Visibility.Visible;
            Main.Visibility = Visibility.Hidden;
            LoadDataGroup();
        }
        /// <summary>
        /// Вывод данных о группах из базы данных в DataGroup через API
        /// </summary>
        public async void LoadDataGroup()
        {
            ResponseModel<List<Group>> responseGroup = await APIHelper.GetDataAsync<List<Group>>("\"\"", "\"\"", "\"\"", typeof(Group));
            if (responseGroup.StatusCode == 201)
            {
                DataGroup.ItemsSource = responseGroup.Data;
            }
        }
        /// <summary>
        /// Добавление данных о группе в базу данных и отображение в DataGroup через API
        /// Проверка на заполения полей
        /// Проверка в базе данных на наличие группы
        /// Обновление данных о группах в DataGroup через API
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        /// <summary>
        /// Изменение данных о группе в базе данных и отображение в DataGroup через API
        /// Проверка на заполения полей
        /// Проверка в базе данных на наличие группы
        /// Обновление данных о группах в DataGroup через API
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        /// <summary>
        /// Удаление данных о группе из базы данных и отображение в DataGroup через API
        /// Обновление данных о группах в DataGroup через API
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        /// <summary>
        /// Выход из панели "Group" 
        /// Открытие главной панели "Main"
        /// Очищение данных в NameGroupBox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExitGroupBtn_Click(object sender, RoutedEventArgs e)
        {
            Group.Visibility = Visibility.Hidden;
            Main.Visibility = Visibility.Visible;
            NameGroupBox.Text = " ";
        }
        ///Предметы
        /// <summary>
        /// Открытие панели "Subject" 
        /// Закрытие главной панели "Main"
        /// Обновление данных о предметах в DataSubject через API
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddSubjectBtn_Click(object sender, RoutedEventArgs e)
        {
            Subject.Visibility = Visibility.Visible;
            Main.Visibility = Visibility.Hidden;
            LoadDataSubject();
        }
        /// <summary>
        /// Обновление данных о предметах в DataSubject через API
        /// </summary>
        public async void LoadDataSubject()
        {
            ResponseModel<List<Subject>> responseSubject = await APIHelper.GetDataAsync<List<Subject>>("\"\"", "\"\"", "\"\"", typeof(Subject));
            if (responseSubject.StatusCode == 201)
            {
                DataSubjcet.ItemsSource = responseSubject.Data;
            }
        }
        /// <summary>
        /// Добавление данных о предмете в базу данных и отображение в DataSubjcet через API
        /// Проверка на заполения полей
        /// Проверка в базе данных на наличие предмета
        /// Обновление данных о предметах в DataSubjcet через API
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        /// <summary>
        /// Удаление данных о предмете из базы данных и отображение в DataSubjcet через API
        /// Проверка на заполения полей
        /// Проверка в базе данных на наличие предмета
        /// Обновление данных о предметах в DataSubjcet через API
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        /// <summary>
        /// Изменение данных о предмете в базе данных и отображение в DataSubjcet через API
        /// Проверка на заполения полей
        /// Проверка в базе данных на наличие предмета
        /// Обновление данных о предметах в DataSubjcet через API
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        /// <summary>
        /// Открытие главной панели "Main" 
        /// Закрытие "Subject"
        /// Очищение данных в NameSubjcetBox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExitSubjcetBtn_Click(object sender, RoutedEventArgs e)
        {
            Subject.Visibility = Visibility.Hidden;
            Main.Visibility = Visibility.Visible;
            NameSubjcetBox.Text = " ";
        }
        ///Преподаватель
        /// <summary>
        /// Открытие панели "Teacher" 
        /// Закрытие главной панели "Main"
        /// Обновление данных о преподавателях в DataTeacher через API
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddTeachertBtn_Click(object sender, RoutedEventArgs e)
        {
            Teacher.Visibility = Visibility.Visible;
            Main.Visibility = Visibility.Hidden;
            LoadDataTeacher();
        }
        /// <summary>
        /// Обновление данных о преподавателях в DataTeacher через API
        /// </summary>
        public async void LoadDataTeacher()
        {
            ResponseModel<List<Teacher>> responseTeacher = await APIHelper.GetDataAsync<List<Teacher>>("\"\"", "\"\"", "\"\"", typeof(Teacher));
            if (responseTeacher.StatusCode == 201)
            {
                DataTeacher.ItemsSource = responseTeacher.Data;
            }
        }
        /// <summary>
        /// Добавление данных о преподавателе в базу данных и отображение в DataTeacher через API
        /// Проверка на заполения полей
        /// Проверка в базе данных на наличие преподавателя
        /// Обновление данных о преподавателях в DataTeacher через API
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        /// <summary>
        /// Изменение данных о преподавателе в базе данных и отображение в DataTeacher через API
        /// Проверка на заполения полей
        /// Проверка в базе данных на наличие преподавателя
        /// Обновление данных о преподавателях в DataTeacher через API
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        /// <summary>
        /// Удаление данных о преподавателе из базы данных и отображение в DataTeacher через API
        /// Проверка на заполения полей
        /// Проверка в базе данных на наличие преподавателя
        /// Обновление данных о преподавателях в DataTeacher через API
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        /// <summary>
        /// Выход из панели "Teacher"
        /// Открытие главной панели "Main"
        /// Очищение данных в NameTeacherBox
        /// Очищение данных в MidlleNameTeacherBox
        /// Очищение данных в FirstNameTeacherBox
        /// Очищение данных в PhoneTeacherBox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExitTeacherBtn_Click(object sender, RoutedEventArgs e)
        {
            Teacher.Visibility = Visibility.Hidden;
            Main.Visibility = Visibility.Visible;
            NameTeacherBox.Text = " ";
            MidlleNameTeacherBox.Text = " ";
            FirstNameTeacherBox.Text = " ";
            PhoneTeacherBox.Text = " ";
        }
        /// <summary>
        /// Открытие метода "excel" из класса "Excel"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExcelBtn_Click(object sender, RoutedEventArgs e)
        {
            Excel excel = new Excel();
            excel.excel();
        }

    }
}
