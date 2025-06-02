using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Hosting;
using Telegram.Bot;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;
using ExcelDataReader;

namespace TelegramExcelBot
{
    public class Student
    {
        public string Id { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string Grade { get; set; }
        public string ClassNum { get; set; }
        public string Gender { get; set; }
        public string Dob { get; set; }
        public string JewishDob { get; set; }
        public string FullAddress { get; set; }
        public string City { get; set; }
        public string SecondAddress { get; set; }
        public string SecondCity { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public string Parent1Id { get; set; }
        public string Parent1Name { get; set; }
        public string Parent1Phone { get; set; }
        public string Parent1Email { get; set; }
        public string Parent2Id { get; set; }
        public string Parent2Name { get; set; }
        public string Parent2Phone { get; set; }
        public string Parent2Email { get; set; }
        public string Major { get; set; }
    }

    public class Teacher
    {
        public string Id { get; set; }
        public string FullNameRaw { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Phone1 { get; set; }
        public string Phone2 { get; set; }
        public string City { get; set; }
        public string FullAddress { get; set; }
        public string Email { get; set; }
        public List<string> Subjects { get; set; } = new List<string>();
        public string Role { get; set; }
    }

    public class SearchConversationState
    {
        public int Step { get; set; }
        public string Option { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public int MenuMessageId { get; set; }
    }

    public class Program
    {
        static TelegramBotClient botClient;
        static Dictionary<long, SearchConversationState> searchStates = new Dictionary<long, SearchConversationState>();
        static List<Student> cachedStudents = new List<Student>();
        static List<Teacher> cachedTeachers = new List<Teacher>();
        static HashSet<long> authorizedUsers = new HashSet<long>();
        const string securityKey = "YafaBish";

        public static async Task Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            // Bind to Render port
            var port = Environment.GetEnvironmentVariable("PORT") ?? "5000";
            builder.WebHost.UseUrls($"http://*:{port}");

            var app = builder.Build();

            // Start the Telegram bot in the background
           // Start the Telegram bot in the background with error handling
_ = Task.Run(async () =>
{
    try
    {
        Console.WriteLine("Starting Telegram bot...");
        await StartBotAsync();
    }
    catch (Exception ex)
    {
        Console.WriteLine("BOT CRASHED: " + ex);
    }
});

// Optional: Self-ping to keep the app warm
_ = Task.Run(async () =>
{
    while (true)
    {
        try
        {
            using var client = new HttpClient();
            var port = Environment.GetEnvironmentVariable("PORT") ?? "5000";
            var result = await client.GetAsync($"http://localhost:{port}/healthz");
            Console.WriteLine("Self-ping: " + result.StatusCode);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Self-ping failed: " + ex.Message);
        }

        await Task.Delay(TimeSpan.FromMinutes(5));
    }
});


            // Minimal HTTP endpoint to bind a port for Render
            app.MapGet("/", () => "Bot is running");
            app.MapGet("/healthz", () => Results.Ok("healthy"));
            await app.RunAsync();
        }

        static async Task StartBotAsync()
        {
            using var mutex = new Mutex(true, "TelegramExcelBotSingleton", out bool createdNew);
            if (!createdNew) return;

            botClient = new TelegramBotClient("7954381826:AAEO7IDqHXd28qeklKXXSXIC-nKzc8G55nU");

            // 1) Delete any existing webhook and drop pending updates
            await botClient.DeleteWebhookAsync(dropPendingUpdates: true);
            // 2) Force-clear the webhook URL so polling won’t conflict
            await botClient.SetWebhookAsync(string.Empty);

            Console.WriteLine("Loading Data...");
            await LoadStudentDataAsync();
            Console.WriteLine("Student Data Loaded, Awaiting Teacher Data...");
            await LoadTeacherDataAsync();
            Console.WriteLine("Teacher Data Loaded");

            var receiverOptions = new ReceiverOptions { AllowedUpdates = Array.Empty<UpdateType>() };
            using var cts = new CancellationTokenSource();

            // 3) Now start long-polling without "Conflict" errors
            botClient.StartReceiving(HandleUpdateAsync, HandleErrorAsync, receiverOptions, cts.Token);

            Console.WriteLine("Bot is running");
            await Task.Delay(Timeout.Infinite);
        }

        static async Task LoadStudentDataAsync()
        {
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            string path = Path.Combine(exeDir, "Data", "ALPON.xlsx");
            if (!File.Exists(path))
                throw new FileNotFoundException($"Could not find student data file at {path}");

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            cachedStudents.Clear();
            using var stream = File.Open(path, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            int row = 0;
            while (reader.Read())
            {
                row++;
                if (row <= 3) continue;
                cachedStudents.Add(new Student
                {
                    Id = reader.GetValue(1)?.ToString() ?? string.Empty,
                    LastName = reader.GetValue(2)?.ToString() ?? string.Empty,
                    FirstName = reader.GetValue(3)?.ToString() ?? string.Empty,
                    Grade = reader.GetValue(4)?.ToString() ?? string.Empty,
                    ClassNum = reader.GetValue(5)?.ToString() ?? string.Empty,
                    Gender = reader.GetValue(6)?.ToString() ?? string.Empty,
                    Dob = reader.GetValue(7)?.ToString() ?? string.Empty,
                    JewishDob = reader.GetValue(8)?.ToString() ?? string.Empty,
                    FullAddress = reader.GetValue(11)?.ToString() ?? string.Empty,
                    City = reader.GetValue(12)?.ToString() ?? string.Empty,
                    SecondAddress = reader.GetValue(13)?.ToString() ?? string.Empty,
                    SecondCity = reader.GetValue(14)?.ToString() ?? string.Empty,
                    Email = reader.GetValue(18)?.ToString() ?? string.Empty,
                    Phone = reader.GetValue(19)?.ToString() ?? string.Empty,
                    Parent1Id = reader.GetValue(20)?.ToString() ?? string.Empty,
                    Parent1Name = reader.GetValue(21)?.ToString() ?? string.Empty,
                    Parent1Phone = reader.GetValue(22)?.ToString() ?? string.Empty,
                    Parent1Email = reader.GetValue(23)?.ToString() ?? string.Empty,
                    Parent2Id = reader.GetValue(24)?.ToString() ?? string.Empty,
                    Parent2Name = reader.GetValue(25)?.ToString() ?? string.Empty,
                    Parent2Phone = reader.GetValue(26)?.ToString() ?? string.Empty,
                    Parent2Email = reader.GetValue(27)?.ToString() ?? string.Empty,
                    Major = reader.GetValue(28)?.ToString() ?? string.Empty
                });
            }
        }

        static async Task LoadTeacherDataAsync()
        {
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            string path = Path.Combine(exeDir, "Data", "TEACHER.xlsx");
            if (!File.Exists(path))
                throw new FileNotFoundException($"Could not find teacher data file at {path}");

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            cachedTeachers.Clear();
            using var stream = File.Open(path, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            int row = 0;
            while (reader.Read())
            {
                row++;
                if (row <= 1) continue;
                var fullName = reader.GetValue(2)?.ToString() ?? string.Empty;
                var parts = fullName.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                var firstName = parts.LastOrDefault() ?? string.Empty;
                var lastName = string.Join(" ", parts.Take(parts.Length - 1));
                var rawSub = reader.GetValue(8)?.ToString() ?? string.Empty;
                var subs = rawSub
                    .Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .Where(s => s.Length > 0)
                    .ToList();

                cachedTeachers.Add(new Teacher
                {
                    Id = reader.GetValue(1)?.ToString() ?? string.Empty,
                    FullNameRaw = fullName,
                    FirstName = firstName,
                    LastName = lastName,
                    Phone1 = reader.GetValue(3)?.ToString() ?? string.Empty,
                    Phone2 = reader.GetValue(4)?.ToString() ?? string.Empty,
                    City = reader.GetValue(5)?.ToString() ?? string.Empty,
                    FullAddress = reader.GetValue(6)?.ToString() ?? string.Empty,
                    Email = reader.GetValue(7)?.ToString() ?? string.Empty,
                    Subjects = subs,
                    Role = reader.GetValue(9)?.ToString() ?? string.Empty
                });
            }
        }

        static async Task HandleUpdateAsync(ITelegramBotClient bot, Update update, CancellationToken ct)
        {
            if (update.Type == UpdateType.CallbackQuery)
            {
                await HandleCallbackQueryAsync(update.CallbackQuery, ct);
                return;
            }
            if (update.Type != UpdateType.Message || update.Message.Text == null) return;

            var msg = update.Message;
            if (!authorizedUsers.Contains(msg.Chat.Id))
            {
                if (msg.Text.StartsWith("/key", StringComparison.OrdinalIgnoreCase))
                {
                    var key = msg.Text.Substring(4).Trim();
                    if (key == securityKey)
                    {
                        authorizedUsers.Add(msg.Chat.Id);
                        await botClient.SendTextMessageAsync(msg.Chat.Id, "האימות הצליח!", cancellationToken: ct);
                    }
                }
                return;
            }

            if (msg.Text.StartsWith("/search", StringComparison.OrdinalIgnoreCase))
            {
                var kb = new InlineKeyboardMarkup(new[]
                {
                    new[] { InlineKeyboardButton.WithCallbackData("תלמיד", "role_student"), InlineKeyboardButton.WithCallbackData("מורה", "role_teacher") }
                });
                await botClient.SendTextMessageAsync(msg.Chat.Id, "בחר סוג חיפוש:", replyMarkup: kb, cancellationToken: ct);
                return;
            }

            if (msg.Text.StartsWith("/start", StringComparison.OrdinalIgnoreCase))
            {
                await botClient.SendTextMessageAsync(msg.Chat.Id, "/search", cancellationToken: ct);
                return;
            }

            if (msg.Text.StartsWith("/help", StringComparison.OrdinalIgnoreCase))
            {
                await botClient.SendTextMessageAsync(msg.Chat.Id, "/search - חיפוש\n/help - עזרה", cancellationToken: ct);
                return;
            }

            if (searchStates.ContainsKey(msg.Chat.Id))
                await ProcessSearchConversation(msg, ct);
        }

        static async Task HandleCallbackQueryAsync(CallbackQuery cq, CancellationToken ct)
        {
            var chatId = cq.Message.Chat.Id;
            await botClient.AnswerCallbackQueryAsync(cq.Id, cancellationToken: ct);

            switch (cq.Data)
            {
                case "go_back_main":
                    var mainKb = new InlineKeyboardMarkup(new[]
                    {
                        new[] { InlineKeyboardButton.WithCallbackData("תלמיד", "role_student"), InlineKeyboardButton.WithCallbackData("מורה", "role_teacher") }
                    });
                    await botClient.EditMessageTextAsync(chatId, cq.Message.MessageId, "בחר סוג חיפוש:", replyMarkup: mainKb, cancellationToken: ct);
                    break;
                case "role_student":
                    var skb = new InlineKeyboardMarkup(new[]
                    {
                        new[] { InlineKeyboardButton.WithCallbackData("תעודת זהות", "student_search_id"), InlineKeyboardButton.WithCallbackData("שם מלא", "student_search_fullname") },
                        new[] { InlineKeyboardButton.WithCallbackData("🔙 חזור", "go_back_main") }
                    });
                    await botClient.EditMessageTextAsync(chatId, cq.Message.MessageId, "בחר חיפוש תלמיד:", replyMarkup: skb, cancellationToken: ct);
                    break;
                case "role_teacher":
                    var tkb = new InlineKeyboardMarkup(new[]
                    {
                        new[] { InlineKeyboardButton.WithCallbackData("תעודת זהות", "teacher_search_id"), InlineKeyboardButton.WithCallbackData("שם מלא", "teacher_search_fullname"), InlineKeyboardButton.WithCallbackData("מקצוע", "teacher_search_subject") },
                        new[] { InlineKeyboardButton.WithCallbackData("🔙 חזור", "go_back_main") }
                    });
                    await botClient.EditMessageTextAsync(chatId, cq.Message.MessageId, "בחר חיפוש מורה:", replyMarkup: tkb, cancellationToken: ct);
                    break;
                case "student_search_id":
                case "student_search_fullname":
                case "teacher_search_id":
                case "teacher_search_fullname":
                case "teacher_search_subject":
                    searchStates[chatId] = new SearchConversationState { Option = cq.Data, Step = 1, MenuMessageId = cq.Message.MessageId };
                    await botClient.DeleteMessageAsync(chatId, cq.Message.MessageId, cancellationToken: ct);
                    string prompt = cq.Data.EndsWith("_id")
                        ? "אנא הזן תעודת זהות:"
                        : cq.Data.EndsWith("_fullname")
                            ? "אנא הזן שם פרטי (או 'ללא'):"
                            : "אנא הזן מקצוע (או 'ללא'):";
                    await botClient.SendTextMessageAsync(chatId, prompt, cancellationToken: ct);
                    break;
            }
        }

        static async Task ProcessSearchConversation(Message message, CancellationToken ct)
        {
            var chatId = message.Chat.Id;
            var state = searchStates[chatId];
            var text = message.Text.Trim();
            bool isTeacher = state.Option.StartsWith("teacher");

            if (state.Option.EndsWith("_fullname") && state.Step == 1)
            {
                state.FirstName = text.Equals("ללא", StringComparison.OrdinalIgnoreCase) ? string.Empty : text;
                state.Step = 2;
                await botClient.SendTextMessageAsync(chatId, "אנא הזן שם משפחה (או 'ללא'):", cancellationToken: ct);
                return;
            }

            string input = text.Equals("ללא", StringComparison.OrdinalIgnoreCase) ? string.Empty : text;
            string fn = state.Option.EndsWith("_fullname") ? state.FirstName : string.Empty;
            string ln = (state.Option.EndsWith("_fullname") && state.Step == 2) ? input : string.Empty;
            string idInput = state.Option.EndsWith("_id") ? input : string.Empty;
            string subjectInput = state.Option == "teacher_search_subject" ? input : string.Empty;

            if (isTeacher)
            {
                fn = fn?.Replace('״', '"');
                ln = ln?.Replace('״', '"');
                subjectInput = subjectInput?.Replace('״', '"');
            }

            if (!isTeacher && state.Option.EndsWith("_fullname") && string.IsNullOrEmpty(fn) && string.IsNullOrEmpty(ln))
            {
                await botClient.SendTextMessageAsync(chatId, "לא ניתן לבצע חיפוש ללא שם.", cancellationToken: ct);
            }
            else if (isTeacher && state.Option.EndsWith("_fullname") && string.IsNullOrEmpty(fn) && string.IsNullOrEmpty(ln) && string.IsNullOrEmpty(subjectInput))
            {
                await botClient.SendTextMessageAsync(chatId, "לא ניתן לבצע חיפוש ללא קריטריון.", cancellationToken: ct);
            }
            else
            {
                if (isTeacher)
                    await PerformTeacherSearch(message, fn, ln, subjectInput, ct);
                else
                    await PerformStudentSearch(message, fn, ln, idInput, ct);
            }

            searchStates.Remove(chatId);
        }

        static async Task PerformStudentSearch(Message message, string fn, string ln, string idInput, CancellationToken ct)
        {
            var results = new List<string>();
            foreach (var s in cachedStudents)
            {
                bool match = !string.IsNullOrEmpty(idInput)
                             ? s.Id.Contains(idInput)
                             : ((string.IsNullOrEmpty(fn) || s.FirstName.Contains(fn, StringComparison.OrdinalIgnoreCase)) &&
                                (string.IsNullOrEmpty(ln) || s.LastName.Contains(ln, StringComparison.OrdinalIgnoreCase)));
                if (!match) continue;

                var sb = new StringBuilder();
                sb.AppendLine("----------------------------------------");
                sb.AppendLine($"תעודת זהות: {s.Id}");
                sb.AppendLine($"שם משפחה: {s.LastName}");
                sb.AppendLine($"שם פרטי: {s.FirstName}");
                sb.AppendLine($"כיתה: {s.Grade} {s.ClassNum}");
                sb.AppendLine($"מין: {s.Gender}");
                sb.AppendLine($"תאריך לידה: {s.Dob}");
                sb.AppendLine($"תאריך לידה עברי: {s.JewishDob}");
                sb.AppendLine($"כתובת מלאה: {s.FullAddress}, {s.City}");
                sb.AppendLine($"כתובת שנייה: {s.SecondAddress}, {s.SecondCity}");
                sb.AppendLine($"אימייל: {s.Email}");
                sb.AppendLine($"טלפון: {s.Phone}");
                sb.AppendLine($"פרטי הורה 1: {s.Parent1Id} - {s.Parent1Name} - {s.Parent1Phone} - {s.Parent1Email}");
                sb.AppendLine($"פרטי הורה 2: {s.Parent2Id} - {s.Parent2Name} - {s.Parent2Phone} - {s.Parent2Email}");
                sb.AppendLine($"מגמה: {s.Major}");
                results.Add(sb.ToString());
            }

            if (results.Count == 0)
            {
                await botClient.SendTextMessageAsync(message.Chat.Id, "לא נמצאו תוצאות.", cancellationToken: ct);
                return;
            }

            var chunk = new StringBuilder();
            foreach (var rec in results)
            {
                if (chunk.Length + rec.Length > 3500)
                {
                    await botClient.SendTextMessageAsync(message.Chat.Id, chunk.ToString(), cancellationToken: ct);
                    chunk.Clear();
                }
                chunk.Append(rec);
            }
            if (chunk.Length > 0)
                await botClient.SendTextMessageAsync(message.Chat.Id, chunk.ToString(), cancellationToken: ct);
        }

        static async Task PerformTeacherSearch(Message message, string fn, string ln, string subjectInput, CancellationToken ct)
        {
            var results = new List<string>();
            foreach (var t in cachedTeachers)
            {
                bool match = ((string.IsNullOrEmpty(fn) || t.FirstName.Contains(fn, StringComparison.OrdinalIgnoreCase)) &&
                              (string.IsNullOrEmpty(ln) || t.LastName.Contains(ln, StringComparison.OrdinalIgnoreCase))) &&
                             (string.IsNullOrEmpty(subjectInput) || t.Subjects.Any(sub => sub.Contains(subjectInput, StringComparison.OrdinalIgnoreCase)));
                if (!match) continue;

                var sb = new StringBuilder();
                sb.AppendLine("----------------------------------------");
                sb.AppendLine($"תעודת זהות: {t.Id}");
                sb.AppendLine($"שם משפחה: {t.LastName}");
                sb.AppendLine($"שם פרטי: {t.FirstName}");
                sb.AppendLine($"מס' טלפון 1: {t.Phone1}");
                sb.AppendLine($"מס' טלפון 2: {t.Phone2}");
                sb.AppendLine($"עיר: {t.City}");
                sb.AppendLine($"כתובת: {t.FullAddress}");
                sb.AppendLine($"מייל: {t.Email}");
                sb.AppendLine($"מקצועות: {string.Join(", ", t.Subjects)}");
                sb.AppendLine($"תפקיד: {t.Role}");
                results.Add(sb.ToString());
            }

            if (results.Count == 0)
            {
                await botClient.SendTextMessageAsync(message.Chat.Id, "לא נמצאו תוצאות.", cancellationToken: ct);
                return;
            }

            var chunk = new StringBuilder();
            foreach (var rec in results)
            {
                if (chunk.Length + rec.Length > 3500)
                {
                    await botClient.SendTextMessageAsync(message.Chat.Id, chunk.ToString(), cancellationToken: ct);
                    chunk.Clear();
                }
                chunk.Append(rec);
            }
            if (chunk.Length > 0)
                await botClient.SendTextMessageAsync(message.Chat.Id, chunk.ToString(), cancellationToken: ct);
        }

        static Task HandleErrorAsync(ITelegramBotClient bot, Exception ex, CancellationToken cancellationToken)
        {
            if (ex is ApiRequestException api && api.Message.Contains("Conflict"))
            {
                Console.WriteLine("Conflict detected.");
                return Task.CompletedTask;
            }
            Console.WriteLine(ex.Message);
            return Task.CompletedTask;
        }
    }
}
