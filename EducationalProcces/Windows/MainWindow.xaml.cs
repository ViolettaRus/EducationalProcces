using EducationalProcces.Windows;
using System.Collections.Generic;
using System.Windows;

namespace EducationalProcces
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            APIHelper.ActivateLink();
            InitializeComponent();
        }

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
