using System;
using System.Diagnostics;
using System.Globalization;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using Calendar;
using Google.Apis.Calendar.v3.Data;

namespace ConsoleCalendar
{
    class Program
    {
        static CalendarManger calender = new CalendarManger();
        static DateTime selectedDate = DateTime.Now;
        static int[] cruntDays = new int[5];
        static Events events;
        static int cursorLength;
        static void Main(string[] args)
        {
            Maximize();
            bool loop = true;
            Drow();
            while (loop)
            {
                string firstLine = String.Format("{0}{1}",
                    "Write (Z) to show Next month.".PadRight(60, ' '),
                    "Write (X) to show past month.");
                cursorLength = firstLine.Length;
                Console.WriteLine("\n");
                Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                Console.WriteLine(firstLine);

                string theredLine = String.Format("{0}{1}",
                    "Write (Y) to move to selected year.".PadRight(60, ' '),
                    "Write (M) to move to selected month.");
                Console.WriteLine("\n");
                Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                Console.WriteLine(theredLine);

                string scoundLine = String.Format("{0}{1}",
                    "Write (D) to day to mange events.".PadRight(60,' '),
                    "Write (Q) to close the app.");
                Console.WriteLine("\n");
                Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                Console.WriteLine(scoundLine);

                Console.WriteLine("\n");
                Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                Console.Write("Enter you choice: ");

                string userChoose = Console.ReadLine();

                userChoose = userChoose.Replace('(', ' ').Replace(')', ' ').Trim().ToLower();

                switch (userChoose)
                {
                    case "z":
                        selectedDate = selectedDate.AddMonths(1);
                        Drow();
                        break;
                    case "x":
                        selectedDate = selectedDate.AddMonths(-1);
                        Drow();
                        break;
                    case "y":
                        MoveToYear();
                        break;
                    case "m":
                        MoveToMonth();
                        break;
                    case "d":
                        SelectedDay();
                        break;
                    case "q":
                        loop = false;
                        break;
                    default:
                        Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                        Console.WriteLine("Worng entry...!");
                        break;
                }
            }
        }

        static void Drow30DaysMonth()
        {
            int monthDay = 1;
            
            for (int z = 0; z < 7; z++)
            {
                Console.SetCursorPosition((Console.WindowWidth - 116) / 2, Console.CursorTop);
                Console.WriteLine(new string('-', 116));
                if (z != 6)
                {
                    for (int x = 0; x < 9; x++)
                    {
                        Console.SetCursorPosition((Console.WindowWidth - 116) / 2, Console.CursorTop);
                        int daysCount = 0;
                        for (int i = 0; i < 6; i++)
                        {
                            if (i == 5)
                            {
                                Console.WriteLine('|');
                                break;
                            }
                            
                            if (x == 0)
                            {
                                
                                if (monthDay < 10)
                                    Console.Write('|' + monthDay.ToString() + new string(' ', 21));
                                else
                                    Console.Write('|' + monthDay.ToString() + new string(' ', 20));
                                cruntDays[daysCount] = monthDay;
                                monthDay++;
                                daysCount++;
                            }
                            else
                            {
                                if (!WriteTheEvent(cruntDays[daysCount]))
                                    Console.Write('|' + new string(' ', 22));
                                daysCount++;
                            }
                        }
                    }

                }
            }
        }

        static void Drow31DaysMonth()
        {
            
            Drow30DaysMonth();
            
            for (int i = 0; i < 9; i++)
            {
                Console.SetCursorPosition((Console.WindowWidth - 116) / 2, Console.CursorTop);
                if (i == 0)
                    Console.WriteLine(new string("|31" + new string(' ', 20) + '|'));
                else
                {
                    if (!WriteTheEvent(31))
                        Console.WriteLine(new string('|' + new string(' ', 22) + '|'));
                }
            }
            Console.SetCursorPosition((Console.WindowWidth - 116) / 2, Console.CursorTop);
            Console.WriteLine(new string('-', 24));
        }

        static void Drow()
        {
            Console.Clear();
            events = calender.GetMonthEvents(selectedDate);
            string head = String.Format("{0} - {1}", selectedDate.ToString("MMMM", CultureInfo.CreateSpecificCulture("en")), selectedDate.Year);
            Console.SetCursorPosition((Console.WindowWidth - head.Length) / 2, Console.CursorTop);
            Console.WriteLine(head);
            if (selectedDate.Month % 2 == 0)
                Drow30DaysMonth();
            else
                Drow31DaysMonth();
        }

        static void MoveToMonth()
        {
            Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
            Console.Write("Enter number btween 1-12 to move to this month: ");
            string usertChoose = Console.ReadLine();
            if (int.TryParse(usertChoose, out int monthNumber) && monthNumber >= 1 && monthNumber <= 12)
                selectedDate = new DateTime(selectedDate.Year, monthNumber, 1);
            Drow();
        }

        static void MoveToYear()
        {
            Console.WriteLine("\n");
            Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
            Console.Write("Enter the year number (MUST BE FULL NUMBER BTWEEN 2000-2050): ");
            string usertChoose = Console.ReadLine();
            if (int.TryParse(usertChoose, out int yearNumber) && yearNumber >= 2000 && yearNumber <= 2050)
                selectedDate = new DateTime(yearNumber, 1, 1);
            Drow();
        }

        static void SelectedDay()
        {
            string userselect;

            Console.WriteLine("\n");
            Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
            Console.Write("Enter day in currnt month to mange: ");
            userselect = Console.ReadLine();

            if (int.TryParse(userselect, out int dayNumber) && dayNumber >= 1 && dayNumber <= DateTime.DaysInMonth(selectedDate.Year, selectedDate.Month))
            {
                selectedDate = new DateTime(selectedDate.Year, selectedDate.Month, dayNumber);
                events = calender.GetDayEvents(selectedDate);
                Console.WriteLine("\n");
                Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                Console.WriteLine("Your selected day is: " + selectedDate.ToString("d"));
                Console.WriteLine("\n");
                Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                Console.WriteLine("You have {0} event/events is this day!", events.Items.Count);
                for (int i = 0; i < events.Items.Count; i++)
                {
                    Console.WriteLine("\n");
                    Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                    Console.WriteLine("{0}- {1}", i + 1, events.Items[i].Summary);
                    Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                    if (!String.IsNullOrEmpty(events.Items[i].Start.DateTime.ToString()))
                        Console.WriteLine("   {0} - {1}", events.Items[i].Start.DateTime.ToString(), events.Items[i].End.DateTime.ToString());
                    else
                        Console.WriteLine("   Event has no time :(");
                    Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                    if (!String.IsNullOrEmpty(events.Items[i].Description))
                        Console.WriteLine("   Event Description: {0}", events.Items[i].Description);
                    else
                        Console.WriteLine("   Event has no Description :(");
                }

                string firstLine = String.Format("{0}{1}",
                    "Write (A) to Add new event.".PadRight(60, ' '),
                    "Write (U) to Update event.");
                cursorLength = firstLine.Length;
                Console.WriteLine("\n");
                Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                Console.WriteLine(firstLine);

                string theredLine = String.Format("{0}{1}",
                    "Write (D) to Delete event.".PadRight(60, ' '),
                    "Write (B) to go back to the main menu.");
                Console.WriteLine("\n");
                Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                Console.WriteLine(theredLine);
                Console.WriteLine("\n");
                Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                Console.Write("Enter you choice: ");
                userselect = Console.ReadLine();

                userselect = userselect.Replace('(', ' ').Replace(')', ' ').Trim().ToLower();
                switch (userselect)
                {
                    case "a":
                        AddEvent();
                        Drow();
                        
                        break;
                    case "u":
                        Console.WriteLine("\n");
                        Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                        Console.Write("Enter event numbur to update: ");
                        userselect = Console.ReadLine();
                        if (int.TryParse(userselect, out int ueventNumber) && ueventNumber >= 1 && ueventNumber <= events.Items.Count)
                        { 
                            Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                            Console.Write("To confirm the update enter (Y): ");
                            userselect = Console.ReadLine();
                            if (userselect.Replace('(',' ').Replace(')', ' ').Trim().ToLower() == "y")
                            {
                                UpdateEvent(events.Items[ueventNumber - 1]);    
                            }
                        }
                        Drow();
                        break;
                    case "d":
                        Console.WriteLine("\n");
                        Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                        Console.Write("Enter event numbur to delete: ");
                        userselect = Console.ReadLine();
                        if (int.TryParse(userselect, out int eventNumber) && eventNumber >= 1 && eventNumber <= events.Items.Count)
                        {
                            Console.WriteLine("\n");
                            Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                            Console.WriteLine("Your about to delete event {0}- {1}: ", eventNumber, events.Items[eventNumber - 1].Summary);
                            Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                            Console.Write("To confirm the delete enter (Y) as it is with the parentheses: ");
                            userselect = Console.ReadLine();
                            if (userselect == "(Y)")
                            {
                                calender.DeleteEvent(events.Items[eventNumber - 1]);
                                SendEmail(events.Items[eventNumber - 1],EventAction.Delete);
                            }
                        }
                        Drow();
                        break;
                    default:
                        Drow();
                        break;
                }
            }
            else
            {
                Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
                Console.WriteLine("Worng Entry.....!");
            }
        }

        enum EventAction
        {
            Add,
            Update,
            Delete
        }

        static void SendEmail(Event eventToMail,EventAction eventAction)
        {
            string emailBody = "";
            string emailSubject = "";
            switch (eventAction)
            {
                case EventAction.Add:
                    emailBody = String.Format("<h1>Hello test email</h1><h2>You add a new event titled: {0}</h2><p>Event date: {1}</p><p>Event discription: {2}</p>",
                        eventToMail.Summary,selectedDate.ToShortDateString(),eventToMail.Description);
                    emailSubject = "Add new Event to calnder";
                    break;
                case EventAction.Update:
                    emailBody = String.Format("<h1>Hello test email</h1><h2>You update an event titled: {0}</h2><p>Event date: {1}</p><p>Event discription: {2}</p>",
                        eventToMail.Summary, selectedDate.ToShortDateString(), eventToMail.Description);
                    emailSubject = "Update Event in calnder";
                    break;
                case EventAction.Delete:
                    emailBody = String.Format("<h1>Hello test email</h1><h2>You deleted an event titled: {0}</h2><p>Event date: {1}</p><p>Event discription: {2}</p>",
                        eventToMail.Summary, selectedDate.ToShortDateString(), eventToMail.Description);
                    emailSubject = "Delete Event in calnder";
                    break;
            }
            Console.WriteLine("\n");
            Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
            Console.Write("Enter your emil adress to send update : ");

            string sendTo = Console.ReadLine();

            if (IsValidEmail(sendTo))
            {
                using MailMessage mail = new MailMessage();
                mail.From = new MailAddress(@"tcalender64@outlook.com");
                mail.To.Add(sendTo);
                mail.Subject = emailSubject;
                mail.Body = emailBody;
                mail.IsBodyHtml = true;

                using SmtpClient smtp = new SmtpClient("smtp-mail.outlook.com", 587);
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new NetworkCredential(@"tcalender64@outlook.com", "Tt0932311301");
                smtp.EnableSsl = true;
                smtp.Send(mail);
            }
        }

        static bool AddEvent()
        {
            Event addEvent = new Event
            {
                Id =  Guid.NewGuid().ToString().Substring(0,7) 
            };

            bool done = false;

            Console.WriteLine("\n");
            Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
            Console.Write("Enter the event name (YOU MUST ENTER IT IT CAN'T BE BLANK): ");
            string eventName = Console.ReadLine();

            Console.WriteLine("\n");
            Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
            Console.Write("Enter the event start time (ex 10:00:00) if you want you can leave it blank: ");
            string eventStartTime = Console.ReadLine();

            Console.WriteLine("\n");
            Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
            Console.Write("Enter the event end time (ex 10:00:00) if you want you can leave it blank: ");
            string eventEndTime = Console.ReadLine();

            Console.WriteLine("\n");
            Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
            Console.Write("Enter the event discription if you want you can leave it blank: ");
            string eventDiscription = Console.ReadLine();

            Console.WriteLine("\n");
            Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
            Console.Write("Enter the event location if you want you can leave it blank: ");
            string eventLocation = Console.ReadLine();

            if (!String.IsNullOrEmpty(eventName))
            {
                addEvent.Summary = eventName;
                if (TimeSpan.TryParse(eventStartTime, out TimeSpan startTime))
                {
                    selectedDate = new DateTime(selectedDate.Year, selectedDate.Month, selectedDate.Day, startTime.Hours, startTime.Minutes, 00);
                    addEvent.Start = new EventDateTime
                    {
                        DateTime = new DateTime(selectedDate.Year, selectedDate.Month, selectedDate.Day, startTime.Hours, startTime.Minutes, 00),
                        TimeZone= "Africa/Cairo"
                    };
                }
                else
                {
                    addEvent.Start = new EventDateTime
                    {
                        DateTime = new DateTime(selectedDate.Year, selectedDate.Month, selectedDate.Day, 00, 00, 00),
                        TimeZone = "Africa/Cairo"
                    };
                }

                if (TimeSpan.TryParse(eventEndTime, out TimeSpan endTime))
                {
                    selectedDate = new DateTime(selectedDate.Year, selectedDate.Month, selectedDate.Day, endTime.Hours, endTime.Minutes, 00);
                    addEvent.End = new EventDateTime 
                    {
                        DateTime = new DateTime(selectedDate.Year, selectedDate.Month, selectedDate.Day, endTime.Hours, endTime.Minutes, 00),
                        TimeZone = "Africa/Cairo"
                    };
                }
                else
                {
                    addEvent.End = new EventDateTime
                    {
                        DateTime = new DateTime(selectedDate.Year, selectedDate.Month, selectedDate.Day, 23, 59, 59),
                        TimeZone = "Africa/Cairo"
                    };
                }

                if (!String.IsNullOrEmpty(eventDiscription))
                    addEvent.Description = eventDiscription;

                if (!String.IsNullOrEmpty(eventLocation))
                    addEvent.Location = eventLocation;

                if (calender.AddNewEvent(addEvent))
                    SendEmail(addEvent, EventAction.Add);
                done = true;
            }
            return done;
        }

        static bool UpdateEvent(Event updateEvent)
        {
            Event updatedEvent = updateEvent;
            bool done = false;

            Console.WriteLine("\n");
            Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
            Console.Write("Enter the event name or keep the old by leave it blank: ");
            string eventName = Console.ReadLine();

            Console.WriteLine("\n");
            Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
            Console.Write("Enter the event start time (ex 10:00:00) if you want you can leave it blank: ");
            string eventStartTime = Console.ReadLine();

            Console.WriteLine("\n");
            Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
            Console.Write("Enter the event end time (ex 10:00:00) if you want you can leave it blank: ");
            string eventEndTime = Console.ReadLine();

            Console.WriteLine("\n");
            Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
            Console.Write("Enter the event discription if you want you can leave it blank: ");
            string eventDiscription = Console.ReadLine();

            Console.WriteLine("\n");
            Console.SetCursorPosition((Console.WindowWidth - cursorLength) / 2, Console.CursorTop);
            Console.Write("Enter the event location if you want you can leave it blank: ");
            string eventLocation = Console.ReadLine();

            if (!String.IsNullOrEmpty(eventName))
                updatedEvent.Summary = eventName;
            if (TimeSpan.TryParse(eventStartTime, out TimeSpan startTime))
            {
                selectedDate = new DateTime(selectedDate.Year, selectedDate.Month, selectedDate.Day, startTime.Hours, startTime.Minutes, 00);
                updatedEvent.Start = new EventDateTime
                {
                    DateTime = selectedDate,
                    TimeZone = "Africa/Cairo"
                };
            }

            if (TimeSpan.TryParse(eventEndTime, out TimeSpan endTime))
            {
                selectedDate = new DateTime(selectedDate.Year, selectedDate.Month, selectedDate.Day, endTime.Hours, endTime.Minutes, 00);
                updatedEvent.End = new EventDateTime
                {
                    DateTime = selectedDate,
                    TimeZone = "Africa/Cairo"
                }; 
            }

            if (!String.IsNullOrEmpty(eventDiscription))
                updatedEvent.Description = eventDiscription;

            if (!String.IsNullOrEmpty(eventLocation))
                updatedEvent.Location = eventLocation;

            if (calender.UpdateEvent(updatedEvent))
                SendEmail(updateEvent, EventAction.Update);
            return done;
        }

        static bool WriteTheEvent(int monthDay)
        {
            bool dayHaveEvent = false;
            Event ev = null;
            foreach (var item in events.Items)
            {
                string when = item.Start.DateTime.ToString();
                if (String.IsNullOrEmpty(when))
                {
                    when = item.Start.Date;
                }
                if (DateTime.Parse(when).Day == monthDay)
                {
                    string datat;
                    if (String.IsNullOrEmpty(item.Start.DateTime.ToString()))
                        datat = String.Format("{0} ({1})", item.Summary, "All day");
                    else
                    {
                        DateTime dateTime = DateTime.Parse(item.Start.DateTime.ToString());
                        datat = String.Format("{0} ({1})", item.Summary, dateTime.ToString("hh:mm tt", CultureInfo.CreateSpecificCulture("en")));
                    }
                    dayHaveEvent = true;
                    
                    if (datat.Length > 22)
                        datat = datat.Substring(0, 22);
                    Console.Write('|' + datat + new string(' ', 22 - datat.Length));
                    if (monthDay == 31)
                        Console.WriteLine('|');
                    ev = item;
                    break;
                }
            }
            if (dayHaveEvent)
                events.Items.Remove(ev);
            return dayHaveEvent;
        }

        static bool IsValidEmail(string email)
        {
            var trimmedEmail = email.Trim();

            if (trimmedEmail.EndsWith("."))
            {
                return false; // suggested by @TK-421
            }
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == trimmedEmail;
            }
            catch
            {
                return false;
            }
        }

        #region Full Screen for Console

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(System.IntPtr hWnd, int cmdShow);

        private static void Maximize()
        {
            Process p = Process.GetCurrentProcess();
            ShowWindow(p.MainWindowHandle, 3);
        }
        #endregion
    }
}