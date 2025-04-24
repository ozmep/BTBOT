using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Telegram.Bot;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;
using ExcelDataReader;

namespace TelegramExcelBot
{
    // Model for a student record.
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
        public string Parent2Email { get; set; }
        public string Major { get; set; }
        public string Parent2Phone { get; set; }
    }

    // Conversation state for /search.
    public class SearchConversationState
    {
        public int Step { get; set; }          // 1: waiting for first input; 2: waiting for second input.
        public string Option { get; set; }       // "id" or "fullname"
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public int MenuMessageId { get; set; }   // The message id of the inline keyboard menu.
    }

    class Program
    {
        static TelegramBotClient botClient;
        // Conversation state per chat.
        static Dictionary<long, SearchConversationState> searchStates = new Dictionary<long, SearchConversationState>();
        // Cached student records.
        static List<Student> cachedStudents = new List<Student>();

        // Security: list of authorized user chat IDs.
        static HashSet<long> authorizedUsers = new HashSet<long>();
        // Set secret key to "YafaBish"
        static readonly string securityKey = "YafaBish";

        static async Task Main(string[] args)
        {
            // Updated token:
            string botToken = "7954381826:AAFo756E7wGIS_k4ur0pSs3TeJTV4QRMjjI";
            botClient = new TelegramBotClient(botToken);

            // Delete any existing webhook to avoid conflicts.
            await botClient.DeleteWebhookAsync();

            // Load the Excel data into memory.
            await LoadStudentDataAsync();

            // Set up polling options (polling via getUpdates).
            var receiverOptions = new ReceiverOptions { AllowedUpdates = Array.Empty<UpdateType>() };
            using var cts = new CancellationTokenSource();

            // Start receiving updates using long polling.
            botClient.StartReceiving(HandleUpdateAsync, HandleErrorAsync, receiverOptions, cts.Token);

            Console.WriteLine("Bot is running with polling.");
            await Task.Delay(Timeout.Infinite);
        }


        // Load Excel data into the cachedStudents list.
        static async Task LoadStudentDataAsync()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "ALPON.xlsx");
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            try
            {
                // Clear any existing data.
                cachedStudents.Clear();

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    int rowIndex = 0;
                    while (reader.Read())
                    {
                        rowIndex++;
                        if (rowIndex <= 3) // Skip the first three rows.
                            continue;

                        var student = new Student
                        {
                            Id = reader.GetValue(1)?.ToString() ?? "",
                            LastName = reader.GetValue(2)?.ToString() ?? "",
                            FirstName = reader.GetValue(3)?.ToString() ?? "",
                            Grade = reader.GetValue(4)?.ToString() ?? "",
                            ClassNum = reader.GetValue(5)?.ToString() ?? "",
                            Gender = reader.GetValue(6)?.ToString() ?? "",
                            Dob = reader.GetValue(7)?.ToString() ?? "",
                            JewishDob = reader.GetValue(8)?.ToString() ?? "",
                            FullAddress = reader.GetValue(11)?.ToString() ?? "",
                            City = reader.GetValue(12)?.ToString() ?? "",
                            SecondAddress = reader.GetValue(13)?.ToString() ?? "",
                            SecondCity = reader.GetValue(14)?.ToString() ?? "",
                            Email = reader.GetValue(18)?.ToString() ?? "",
                            Phone = reader.GetValue(19)?.ToString() ?? "",
                            Parent1Id = reader.GetValue(20)?.ToString() ?? "",
                            Parent1Name = reader.GetValue(21)?.ToString() ?? "",
                            Parent1Phone = reader.GetValue(22)?.ToString() ?? "",
                            Parent1Email = reader.GetValue(23)?.ToString() ?? "",
                            Parent2Id = reader.GetValue(24)?.ToString() ?? "",
                            Parent2Name = reader.GetValue(25)?.ToString() ?? "",
                            Parent2Phone = reader.GetValue(26)?.ToString() ?? "",
                            Parent2Email = reader.GetValue(27)?.ToString() ?? "",
                            Major = reader.GetValue(28)?.ToString() ?? ""
                        };

                        cachedStudents.Add(student);
                    }
                }

                Console.WriteLine($"Loaded {cachedStudents.Count} student records.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error loading student data: " + ex.Message);
            }

            await Task.CompletedTask;
        }

        // Main update handler.
        static async Task HandleUpdateAsync(ITelegramBotClient bot, Update update, CancellationToken cancellationToken)
        {
            // Process callback queries (for inline keyboard).
            if (update.Type == UpdateType.CallbackQuery)
            {
                await HandleCallbackQueryAsync(update.CallbackQuery, cancellationToken);
                return;
            }

            if (update.Type != UpdateType.Message)
                return;

            var message = update.Message;
            if (message.Text == null)
                return;

            Console.WriteLine($"Received message from {message.From.FirstName}: {message.Text}");

            // Security barrier: Only process messages from authorized users.
            if (!authorizedUsers.Contains(message.Chat.Id))
            {
                // Check if the user is trying to send the key.
                if (message.Text.StartsWith("/key", StringComparison.OrdinalIgnoreCase))
                {
                    string providedKey = message.Text.Substring(4).Trim();
                    if (providedKey == securityKey)
                    {
                        authorizedUsers.Add(message.Chat.Id);
                        await botClient.SendTextMessageAsync(message.Chat.Id, "האימות הצליח. כעת ניתן להשתמש בבוט.", cancellationToken: cancellationToken);
                    }
                }
                // If not sending a key, simply ignore.
                return;
            }

            // Process conversation if in progress, otherwise process commands.
            if (searchStates.ContainsKey(message.Chat.Id))
            {
                await ProcessSearchConversation(message, cancellationToken);
            }
            else
            {
                await ProcessCommand(message, cancellationToken);
            }
        }

        // Handle inline keyboard callback queries.
        static async Task HandleCallbackQueryAsync(CallbackQuery callbackQuery, CancellationToken cancellationToken)
        {
            var chatId = callbackQuery.Message.Chat.Id;
            if (callbackQuery.Data == "search_id" || callbackQuery.Data == "search_fullname")
            {
                var state = new SearchConversationState();
                state.Option = (callbackQuery.Data == "search_id") ? "id" : "fullname";
                state.Step = 1;
                state.MenuMessageId = callbackQuery.Message.MessageId;
                searchStates[chatId] = state;

                // Delete the inline keyboard message.
                await botClient.DeleteMessageAsync(chatId, state.MenuMessageId, cancellationToken);

                // Prompt for input.
                if (state.Option == "id")
                {
                    await botClient.SendTextMessageAsync(chatId, "אנא הזן תעודת זהות:", cancellationToken: cancellationToken);
                }
                else
                {
                    await botClient.SendTextMessageAsync(chatId, "אנא הזן שם פרטי (או 'ללא' לדילוג):", cancellationToken: cancellationToken);
                }

                await botClient.AnswerCallbackQueryAsync(callbackQuery.Id, cancellationToken: cancellationToken);
            }
        }

        // Process non-search commands.
        static async Task ProcessCommand(Message message, CancellationToken cancellationToken)
        {
            string text = message.Text.Trim();

            if (text.StartsWith("/search", StringComparison.OrdinalIgnoreCase))
            {
                // Show inline keyboard menu.
                var inlineKeyboard = new InlineKeyboardMarkup(new[]
                {
                    new []
                    {
                        InlineKeyboardButton.WithCallbackData("חיפוש לפי תעודת זהות", "search_id"),
                        InlineKeyboardButton.WithCallbackData("חיפוש לפי שם מלא", "search_fullname")
                    }
                });
                await botClient.SendTextMessageAsync(
                    chatId: message.Chat.Id,
                    text: "בחר את אופציית החיפוש:",
                    replyMarkup: inlineKeyboard,
                    cancellationToken: cancellationToken
                );
            }
            else if (text.StartsWith("/help", StringComparison.OrdinalIgnoreCase))
            {
                string helpText = "פקודות זמינות:\n/help - עזרה\n/search - חיפוש תלמיד לפי תעודת זהות או שם מלא.";
                await botClient.SendTextMessageAsync(message.Chat.Id, helpText, cancellationToken: cancellationToken);
            }
            else if (text.StartsWith("/start", StringComparison.OrdinalIgnoreCase))
            {
                await botClient.SendTextMessageAsync(message.Chat.Id, "ברוכים הבאים לבוט חיפוש התלמידים!", cancellationToken: cancellationToken);
            }
            else
            {
                await botClient.SendTextMessageAsync(message.Chat.Id, "פקודה לא מוכרת. הקלד /help לעזרה.", cancellationToken: cancellationToken);
            }
        }

        // Process conversation when in a search session.
        static async Task ProcessSearchConversation(Message message, CancellationToken cancellationToken)
        {
            var chatId = message.Chat.Id;
            var state = searchStates[chatId];
            string text = message.Text.Trim();

            if (state.Option == "id")
            {
                // Use the entered text as ID.
                await PerformSearch(message, "", "", text, cancellationToken);
                searchStates.Remove(chatId);
            }
            else if (state.Option == "fullname")
            {
                if (state.Step == 1)
                {
                    // First input: first name.
                    state.FirstName = text.Equals("ללא", StringComparison.OrdinalIgnoreCase) ? "" : text;
                    state.Step = 2;
                    await botClient.SendTextMessageAsync(chatId, "אנא הזן שם משפחה (או 'ללא' לדילוג):", cancellationToken: cancellationToken);
                }
                else if (state.Step == 2)
                {
                    state.LastName = text.Equals("ללא", StringComparison.OrdinalIgnoreCase) ? "" : text;
                    await PerformSearch(message, state.FirstName, state.LastName, "", cancellationToken);
                    searchStates.Remove(chatId);
                }
            }
        }

        // Perform the search using cached student data, update the "searching" message, and send the final results.
        static async Task PerformSearch(Message message, string firstName, string lastName, string searchId, CancellationToken cancellationToken)
        {
            // Prevent empty searches (all fields empty).
            if (string.IsNullOrEmpty(firstName) && string.IsNullOrEmpty(lastName) && string.IsNullOrEmpty(searchId))
            {
                await botClient.SendTextMessageAsync(message.Chat.Id, "לא ניתן לבצע חיפוש ללא קלט. אנא הזן תעודת זהות או שם חלקי.", cancellationToken: cancellationToken);
                return;
            }

            // Send initial "searching" message.
            var searchingMsg = await botClient.SendTextMessageAsync(
                chatId: message.Chat.Id,
                text: "מחפש, נמצאו 0 תוצאות...",
                cancellationToken: cancellationToken
            );

            StringBuilder resultBuilder = new StringBuilder();
            int countMatches = 0;

            // Search through the cached list.
            foreach (var student in cachedStudents)
            {
                bool match = false;
                if (!string.IsNullOrEmpty(searchId))
                {
                    match = student.Id.Contains(searchId);
                }
                else
                {
                    bool matchFirst = string.IsNullOrEmpty(firstName) || student.FirstName.ToLower().Contains(firstName.ToLower());
                    bool matchLast = string.IsNullOrEmpty(lastName) || student.LastName.ToLower().Contains(lastName.ToLower());
                    match = matchFirst && matchLast;
                }

                if (match)
                {
                    countMatches++;
                    resultBuilder.AppendLine($"תעודת זהות: {student.Id}");
                    resultBuilder.AppendLine($"שם משפחה: {student.LastName}");
                    resultBuilder.AppendLine($"שם פרטי: {student.FirstName}");
                    resultBuilder.AppendLine($"כיתה: {student.Grade} {student.ClassNum}");
                    resultBuilder.AppendLine($"מין: {student.Gender}");
                    resultBuilder.AppendLine($"תאריך לידה: {student.Dob}");
                    resultBuilder.AppendLine($"תאריך לידה עברי: {student.JewishDob}");
                    resultBuilder.AppendLine($"כתובת מלאה: {student.FullAddress}, {student.City}");
                    resultBuilder.AppendLine($"כתובת שנייה: {student.SecondAddress}, {student.SecondCity}");
                    resultBuilder.AppendLine($"אימייל: {student.Email}");
                    resultBuilder.AppendLine($"טלפון: {student.Phone}");
                    resultBuilder.AppendLine($"פרטי הורה 1: {student.Parent1Id} - {student.Parent1Name} - {student.Parent1Phone} - {student.Parent1Email}");
                    resultBuilder.AppendLine($"פרטי הורה 2: {student.Parent2Id} - {student.Parent2Name} - {student.Parent2Phone} - {student.Parent2Email}");
                    resultBuilder.AppendLine($"מגמה: {student.Major}");
                    resultBuilder.AppendLine(new string('-', 40));

                    // Update the searching message every 10 matches.
                    if (countMatches % 10 == 0)
                    {
                        await botClient.EditMessageTextAsync(
                            chatId: message.Chat.Id,
                            messageId: searchingMsg.MessageId,
                            text: $"מחפש, נמצאו בנתיים {countMatches} תוצאות...",
                            cancellationToken: cancellationToken
                        );
                    }
                }
            }

            // Final update.
            await botClient.EditMessageTextAsync(
                chatId: message.Chat.Id,
                messageId: searchingMsg.MessageId,
                text: $"מחפש, נמצאו בנתיים {countMatches} תוצאות...",
                cancellationToken: cancellationToken
            );

            // Delete the searching message.
            await botClient.DeleteMessageAsync(message.Chat.Id, searchingMsg.MessageId, cancellationToken: cancellationToken);

            // Prepare the final result.
            string finalResult = resultBuilder.Length > 0 ? "תוצאות חיפוש:\n" + resultBuilder.ToString() : "לא נמצאו תוצאות.";

            // If the message is too long, split it into chunks (max 4000 characters each).
            const int maxChunkSize = 4000;
            if (finalResult.Length > maxChunkSize)
            {
                for (int i = 0; i < finalResult.Length; i += maxChunkSize)
                {
                    string chunk = finalResult.Substring(i, Math.Min(maxChunkSize, finalResult.Length - i));
                    await botClient.SendTextMessageAsync(message.Chat.Id, chunk, cancellationToken: cancellationToken);
                }
            }
            else
            {
                await botClient.SendTextMessageAsync(message.Chat.Id, finalResult, cancellationToken: cancellationToken);
            }
        }

        static Task HandleErrorAsync(ITelegramBotClient bot, Exception exception, CancellationToken cancellationToken)
        {
            Console.WriteLine($"Error: {exception.Message}");
            return Task.CompletedTask;
        }
    }
}
