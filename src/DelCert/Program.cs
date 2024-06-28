/*
	Простая консольная программа для установки сертификатов с автоматическим подтверждением всплывающих окон безопасности.

	Данная программа является свободным программным обеспечением, распространяющимся по лицензии MIT.
	Копия лицензии: https://opensource.org/licenses/MIT

	Copyright (c) 2024 Otto
	Автор: Otto
	Версия: 28.06.24
	GitHub страница:  https://github.com/Otto17/DelCert
	GitFlic страница: https://gitflic.ru/project/otto/delcert

	г. Омск 2024
*/

using System;                                           // Библиотека предоставляет доступ к базовым классам и функциональности .NET Framework
using System.Runtime.InteropServices;                   // Библиотека предоставляет средства для взаимодействия с нативным кодом в C#
using System.Security.Cryptography.X509Certificates;    // Библиотека предоставляет возможности для работы с сертификатами X.509
using System.Threading;                                 // Библиотека предоставляет средства для работы с потоками исполнения в C#

namespace DelCert
{
    class Program
    {
        //Импорт необходимых функций WinAPI из библиотеки "user32.dll" (для работы с системными всплывающими окнами)
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);   // Используем метод "FindWindow()" для поиска окна по заданным параметрам

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetForegroundWindow(IntPtr hWnd);    // Используем метод "SetForegroundWindow()" для установки указанного окна на передний план

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);   // Используем метод "PostMessage()" для отправки сообщения указанному окну

        [DllImport("user32.dll")]
        static extern void BlockInput(bool fBlockIt);   // Используем метод "BlockInput()" для блокировки или разблокировки ввода пользователей (мышь/клавиатура)

        //Константы событий для клавиш
        private const uint WM_KEYDOWN = 0x0100; // Событие нажатия клавиши
        private const uint WM_KEYUP = 0x0101;   // Событие отпускания клавиши
        private const int VK_TAB = 0x09;        // Клавиша "Tab" на клавиатуре
        private const int VK_RETURN = 0x0D;     // Клавиша "Enter" на клавиатуре

        private static bool keepRunning = true; // Флаг для запуска/остановки нового потока


        static void Main(string[] args)
        {
            if (args.Length < 3)    // Если получили меньше 3-х аргументов, тогда выводим справку
            {
                //Вывод справки
                Console.ForegroundColor = ConsoleColor.White; // Устанавливаем белый цвет для строк ниже
                Console.WriteLine("Использование:");
                Console.WriteLine("DelCert.exe <CurrentUser или LocalMachine> <Название хранилища> <Сертификат_1> <Сертификат_2> ... <Сертификат_N>\n");
                Console.ResetColor(); // Сбрасываем цвет на стандартный

                Console.ForegroundColor = ConsoleColor.DarkGreen; // Устанавливаем тёмно-зелёный цвет для строк ниже
                Console.WriteLine("Для удаления сертификата можно использовать как общее имя \"CN\" (Common Name), так и серийный номер сертификата.\n");
                Console.WriteLine("Поддерживается автоматическое подтверждение всплывающих окон при установке сертификата в \"Root\" текущего пользователя.");
                Console.WriteLine("Для корректного подтверждения всех всплывающих окон производится блокировка мыши и клавиатуры.\n");

                Console.ForegroundColor = ConsoleColor.Blue; // Устанавливаем синий цвет для строк ниже
                Console.WriteLine("Примеры:");
                Console.WriteLine("DelCert \"CurrentUser\" \"My\" \"TEST1\"");
                Console.WriteLine("DelCert \"LocalMachine\" \"Root\" \"3b6300d6d6523e8d96738b3587303438cbf94936\"");
                Console.WriteLine("DelCert \"LocalMachine\" \"CA\" \"TEST3\" \"TEST4\"");
                Console.WriteLine("DelCert \"CurrentUser\" \"TrustedPeople\" \"TEST2\" \"46f1252342e07172fcd2ee27f0083b12ab99d68d\" \"TEST6\"\n");

                Console.ForegroundColor = ConsoleColor.Red; // Устанавливаем красный цвет для строки ниже
                Console.WriteLine("Для блокировки мыши и клавиатуры (во время подтверждения всплывающих окон) программа должна быть запущена с правами администратора!\n");
                Console.ResetColor(); // Сбрасываем цвет на стандартный

                Console.WriteLine("Список названий хранилищ сертификатов:");
                Console.WriteLine("\"My\"               - Личные");
                Console.WriteLine("\"Root\"             - Доверенные корневые центры сертификации");
                Console.WriteLine("\"Trust\"            - Доверительные отношения в предприятии");
                Console.WriteLine("\"CA\"               - Промежуточные центры сертификации");
                Console.WriteLine("\"TrustedPublisher\" - Доверенные издатели");
                Console.WriteLine("\"AuthRoot\"         - Сторонние корневые центры сертификации");
                Console.WriteLine("\"TrustedPeople\"    - Доверенные лица");
                Console.WriteLine("\"AddressBook\"      - Другие пользователи\n");

                Console.ForegroundColor = ConsoleColor.Yellow; // Устанавливаем жёлтый цвет для строк ниже
                Console.WriteLine("Автор Otto, г.Омск 2024");
                Console.WriteLine("GitHub страница:  https://github.com/Otto17/DelCert");
                Console.WriteLine("GitFlic страница: https://gitflic.ru/project/otto/delcert");
                Console.ResetColor(); // Сбрасываем цвет на стандартный
                return;
            }

            //Массивы для получения аргументов
            string location = args[0];                                  // Получаем первый аргумент
            string storeName = args[1];                                 // Получаем второй аргумент
            string[] certIdentifiers = new string[args.Length - 2];     // Создаём массив строк, для получения аргументов после второго
            Array.Copy(args, 2, certIdentifiers, 0, args.Length - 2);   // Копируем все аргументы после второго в "certIdentifiers"

            StoreLocation storeLocation;    // Переменная для выбора хранилища установки сертификата

            //Выбор хранилища в зависимости от указания первого аргумента
            if (location.Equals("CurrentUser", StringComparison.OrdinalIgnoreCase))
            {
                storeLocation = StoreLocation.CurrentUser;
            }
            else if (location.Equals("LocalMachine", StringComparison.OrdinalIgnoreCase))
            {
                storeLocation = StoreLocation.LocalMachine;
            }
            else
            {
                Console.WriteLine("Ошибка: неверное значение для параметра, выберите \"CurrentUser\" или \"LocalMachine\".");   // Выводим ошибку и завершаем работу программы
                return;
            }

            try
            {
                //Создание и запуск нового потока для автоматического подтверждения диалогового окна
                Thread promptHandlerThread = new Thread(AutoConfirmDialog);
                promptHandlerThread.Start();

                //Вызываем метод для удаления сертификата
                foreach (string certIdentifier in certIdentifiers)
                {
                    RemoveCertificate(storeName, storeLocation, certIdentifier);
                }

                //Остановка потока подтверждения
                keepRunning = false;        // Опускаем флаг для остановки потока
                promptHandlerThread.Join(); // Ожидаем завершения работы потока
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при удалении сертификата: {ex.Message}");
                keepRunning = false;    // Опускаем флаг для остановки потока
            }
        }


        //Функция удаления сертификатов
        private static void RemoveCertificate(string storeName, StoreLocation storeLocation, string certIdentifier)
        {
            using (X509Store store = new X509Store(storeName, storeLocation))   //Метод для открытия хранилища сертификатов
            {
                store.Open(OpenFlags.ReadWrite);    // Открываем хранилище для чтения и записи

                //Поиск сертификатов по "CN" (Common Name)
                X509Certificate2Collection certCollection = new X509Certificate2Collection();   // Создаём коллекцию "X509Certificate2Collection" для хранения найденных сертификатов
                foreach (X509Certificate2 cert in store.Certificates)   // для каждого сертификата  в хранилище проверяем его "CN"
                {
                    if (cert.GetNameInfo(X509NameType.SimpleName, false).Equals(certIdentifier, StringComparison.OrdinalIgnoreCase))    // Если совпадает с сертификатом "certIdentifier" (игнорируя регистр) 
                    {
                        certCollection.Add(cert);   // Добавляем сертификат в коллекцию
                    }
                }

                //Если по "CN" не найдено, ищем по серийному номеру
                if (certCollection.Count == 0)
                {
                    certCollection = store.Certificates.Find(X509FindType.FindBySerialNumber, certIdentifier, false);
                }

                //Если не нашли по "CN" и серийному номеру
                if (certCollection.Count == 0)
                {
                    Console.WriteLine($"Сертификат с именем или серийным номером \"{certIdentifier}\" не найден.");
                }
                else
                {
                    foreach (X509Certificate2 cert in certCollection)   // Иначе, для каждого найденного сертификата из "certCollection" выполняем удаление из хранилища
                    {
                        store.Remove(cert);
                        Console.WriteLine($"Сертификат \"{cert.Subject}\" с серийным номером \"{cert.SerialNumber}\" удален из хранилища \"{storeName}\" в расположении \"{storeLocation}\".");
                    }
                }

                store.Close();  // Закрываем хранилище
            }
        }


        //Поиск всплывающего окна для подтверждения удаления сертификатов
        private static IntPtr FindSecurityWarningWindow()
        {
            IntPtr hwnd = FindWindow(null, "Корневое хранилище сертификатов");  // Ищем окно по заголовку
            if (hwnd == IntPtr.Zero)                                            // Если окно не было найдено
            {
                hwnd = FindWindow(null, "Root Certificate Store");  // Повторяем поиск, но уже на Английском языке (если система не Русифицирована)
            }
            return hwnd;
        }


        //Автоматическое подтверждение при удалении сертификатов во всех всплывающих окнах (независимо сколько их будет)
        private static void AutoConfirmDialog()
        {
            try
            {
                BlockInput(true);   // Блокируем ввод с клавиатуры и мыши

                while (keepRunning) // Если новый поток запущен
                {
                    IntPtr hwnd = FindSecurityWarningWindow();  // Ищем окна из данного метода
                    if (hwnd != IntPtr.Zero)                    // Если окно нашли
                    {
                        Console.WriteLine("Найдено всплывающее диалоговое окно. Блокируем ввод и подтверждаем...");
                        SetForegroundWindow(hwnd);  // Помечаем найденное окно как активное

                        //Отправка клавиши "Tab" для перехода на кнопку "ДА" в диалоговом окне
                        PostMessage(hwnd, WM_KEYDOWN, (IntPtr)VK_TAB, IntPtr.Zero);
                        PostMessage(hwnd, WM_KEYUP, (IntPtr)VK_TAB, IntPtr.Zero);

                        Thread.Sleep(100);  // Небольшая задержка для обработки действия

                        //Отправка клавиши "Enter" для подтверждения диалогового окна
                        PostMessage(hwnd, WM_KEYDOWN, (IntPtr)VK_RETURN, IntPtr.Zero);
                        PostMessage(hwnd, WM_KEYUP, (IntPtr)VK_RETURN, IntPtr.Zero);

                        Thread.Sleep(300);  // Небольшая задержка для обработки действия
                    }
                    Thread.Sleep(400);  // Проверка цикла на новые окна каждые 400 мс, пока "RemoveCertificate()" не завершит свою работу по установке цепочки сертификатов
                }
            }
            finally
            {
                BlockInput(false);  // Разблокируем ввод с клавиатуры и мыши
            }
        }
    }
}
