using EducationalProcces.Windows;
using System.Collections.Generic;
using System.Windows;
using System;

namespace EducationalProcces
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// Обращение к методу "ActivateLink" из класса "APIHelper"
        /// </summary>
        public MainWindow()
        {
            APIHelper.ActivateLink();
            InitializeComponent();
        }
        /// <summary>
        /// Проверка на заполение полей
        /// Обращение к методу "Auth"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AuthBtn_Click(object sender, RoutedEventArgs e)
        {

            if (string.IsNullOrEmpty(LoginBox.Text) || string.IsNullOrEmpty(PasswordTextBox.Password))
            {
                MessageBox.Show("Заполните логин или пароль", "Ошибка авторизации",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                Auth();
            }

        }
        /// <summary>
        /// Проверка на правильность введенных данных в Login и Password с данными в базе данных
        /// Переход на окно "Главная"
        /// </summary>
        public async void Auth()
        {
            User user = new User();
            ResponseModel<List<User>> responsUser = await APIHelper.GetDataAsync<List<User>>(nameof(user.Login), LoginBox.Text, "\"\"", typeof(User));

            if (responsUser.StatusCode == 201)
            {
                if (responsUser.Data[0].Password == PasswordTextBox.Password)
                {
                    user = responsUser.Data[0];

                    if (user.Role.ID_Role == 1)
                    {
                        AddWindow addWindow = new();
                        addWindow.Show();
                        this.Close();
                    }
                    else if (user.Role.ID_Role == 2)
                    {
                        Microsoft.Office.Interop.Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
                        ex.Visible = true;
                        ex.Workbooks.Open(@"C:\Расписание1.xlsx", Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }
                }
                else
                {
                    MessageBox.Show("Неверный пароль", "Ошибка авторизации",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show("Такого пользователя не существует", "Ошибка авторизации",
                   MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
