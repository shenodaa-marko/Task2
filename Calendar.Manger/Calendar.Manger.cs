using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using System;
using System.IO;

namespace Calendar
{
    public class CalendarManger
    {
        #region Private Members

        private const string calendarJsonFile = "gc.json";

        private const string calendarID = @"tcalender64@gmail.com";

        private CalendarService calendarService;

        private Events eventsList;
        #endregion

        #region Constractor

        public CalendarManger()
        {
            ReadCalendarCredential();
        }
        #endregion

        #region Private Methods

        private void ReadCalendarCredential()
        {
            using var stream =
                new FileStream(calendarJsonFile, FileMode.Open, FileAccess.Read);
            var confg = Google.Apis.Json.NewtonsoftJsonSerializer.Instance.Deserialize<JsonCredentialParameters>(stream);
            var credential = new ServiceAccountCredential(
               new ServiceAccountCredential.Initializer(confg.ClientEmail)
               {
                   Scopes = new[] { CalendarService.Scope.Calendar }
               }.FromPrivateKey(confg.PrivateKey));

            calendarService = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Calender Test"
            });
        }

        private void GetEvents(DateTime startDate,DateTime endDate)
        {
            if (eventsList != null)
                eventsList.Items.Clear();
            EventsResource.ListRequest listRequest = calendarService.Events.List(calendarID);
            listRequest.TimeMin = startDate;
            listRequest.TimeMax = endDate;
            listRequest.ShowDeleted = false;
            listRequest.SingleEvents = true;
            listRequest.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;
            eventsList = listRequest.Execute();
        }
        #endregion

        #region Public Get Events Methods

        public Events GetMonthEvents(DateTime selectedMonth)
        {
            DateTime startDate = new DateTime(selectedMonth.Year, selectedMonth.Month, 1);
            DateTime endDate = startDate.AddMonths(1);
            GetEvents(startDate, endDate);
            return eventsList;
        }

        public Events GetDayEvents(DateTime selectedDay)
        {
            DateTime startDate = new DateTime(selectedDay.Year, selectedDay.Month, selectedDay.Day,00,00,00);
            DateTime endDate = new DateTime(selectedDay.Year, selectedDay.Month, selectedDay.Day,23,59,59);
            GetEvents(startDate, endDate);
            return eventsList;
        }
        #endregion

        #region Public Add, Update or Delete Event Methods

        public bool AddNewEvent(Event newEvent)
        {
            var InsertRequest = calendarService.Events.Insert(newEvent, calendarID);

            try
            {
                InsertRequest.Execute();
                return true;
            }
            catch (Exception ex)
            {
                string e = ex.Message;
                return false;
            }
        }

        public bool UpdateEvent(Event myEvent)
        { 
            try
            {
                calendarService.Events.Update(myEvent, calendarID, myEvent.Id).Execute();
                return true;
            }
            catch (Exception ex)
            {
                string e = ex.Message;
                return false;
            }
        }

        public bool DeleteEvent(Event myEvent)
        {
            try
            {
                calendarService.Events.Delete(calendarID, myEvent.Id).Execute();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        #endregion
    }
}