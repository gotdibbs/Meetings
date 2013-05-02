﻿using MahApps.Metro.Controls;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using UpcomingMeetings.Model;
using System.Runtime.InteropServices;
using ThreadTask = System.Threading.Tasks.Task;
using System.Windows.Interop;

namespace UpcomingMeetings
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        #region Private

        private ExchangeService _service;

        private string _emailAddress;

        private ExchangeVersion _exchangeVersion;

        private UpcomingMeetings.Properties.Settings _settings = Properties.Settings.Default;

        #endregion

        public MainWindow()
        {
            InitializeComponent();

            // Default control values
            Exchange.SelectedItem = ExchangeVersion.Exchange2010_SP2;
            Email.Text = string.Empty;

            LoadSettings();
        }

        /// <summary>
        /// Checks if we're already setup, if not, loads the settings frame
        /// </summary>
        public void LoadSettings()
        {
            if (!_settings.IsSetup)
            {
                Exchange.ItemsSource = Enum.GetValues(typeof(ExchangeVersion));
                Settings.Visibility = System.Windows.Visibility.Visible;
                UpcomingMeetings.Visibility = System.Windows.Visibility.Hidden;
            }
            else
            {
                _service = null;
                _emailAddress = _settings.Email;
                _exchangeVersion = _settings.ExchangeVersion;
                Settings.Visibility = System.Windows.Visibility.Collapsed;
                UpcomingMeetings.Visibility = System.Windows.Visibility.Visible;
                StartPoll();
            }
        }

        /// <summary>
        /// Kicks off background polling for meeting refresh
        /// </summary>
        public void StartPoll()
        {
            ThreadTask.Run(() => PollForMeetings());
        }

        /// <summary>
        /// Updates list of meetings with a Lync Meeting Uri
        /// </summary>
        public void PollForMeetings()
        {
            if (_service == null)
            {
                try
                {
                    _service = new ExchangeService(_exchangeVersion);
                    _service.UseDefaultCredentials = true;
                    _service.AutodiscoverUrl(_emailAddress);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("A problem was encountered while attempting to connect to Exchange Web Services. Please check network connectivity and settings.");
                    _service = null;
                }
            }

            // Grab calendar folder
            var calendarFolder = new FolderId(WellKnownFolderName.Calendar);
            // Create calendar view (appointment starting or ending between now and 12 hours from now)
            var calendarView = new CalendarView(DateTime.Now, DateTime.Now.AddHours(12));

            // Get Ids of second-class Appointment properties
            var UCOpenedConferenceID = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings, "UCOpenedConferenceID", MapiPropertyType.String);
            var OnlineMeetingExternalLink = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings, "OnlineMeetingExternalLink", MapiPropertyType.String);

            // Definte column set to retrieve from appointment objects
            PropertySet iDPropertySet = new PropertySet(BasePropertySet.FirstClassProperties) { UCOpenedConferenceID };
            calendarView.PropertySet = iDPropertySet;

            // Retrieve appointments and parse for those with Lync conferences
            var lyncMeetings = new List<Appointment>();
            var appointmentResult = _service.FindAppointments(calendarFolder, calendarView);
            foreach (var appointment in appointmentResult)
            {
                object UCconfId = null;
                if(appointment.TryGetProperty(UCOpenedConferenceID, out UCconfId))
                    lyncMeetings.Add(appointment);
            }

            var uris = new List<LocalAppointment>();

            if (lyncMeetings != null && lyncMeetings.Count > 0) {
                // Start get the details of each appointment
                var detailPropertySet = new PropertySet(BasePropertySet.FirstClassProperties) { OnlineMeetingExternalLink };
                _service.LoadPropertiesForItems(lyncMeetings, detailPropertySet);

                // Parse appointments to local model
                
                foreach (Appointment appointment in lyncMeetings)
                {
                    string lyncMeetingUrl = null;
                    appointment.TryGetProperty(OnlineMeetingExternalLink, out lyncMeetingUrl);
                    uris.Add(new LocalAppointment 
                    { 
                        Subject = appointment.Subject, 
                        Location = appointment.Location,
                        Uri = lyncMeetingUrl, 
                        StartTime = appointment.Start 
                    });
                }
            }

            // Update the UI
            Dispatcher.Invoke(() =>
            {
                UpcomingMeetings.ItemsSource = uris;
            });

            // Poll every 5 min
            Thread.Sleep(new TimeSpan(0, 5, 0));
            PollForMeetings();
        }

        #region Event Handlers

        /// <summary>
        /// Handles the click of a join meeting button from the list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LyncUri_Click(object sender, RoutedEventArgs e)
        {
            var item = ((Button)sender).DataContext as LocalAppointment;

            System.Diagnostics.Process.Start("conf:sip:" + item.Uri);
        }

        /// <summary>
        /// Handles save of settings
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveSettings_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(Email.Text))
            {
                MessageBox.Show("Please enter an email address.");
                return;
            }
            if (Exchange.SelectedItem == null)
            {
                MessageBox.Show("Please enter an exchange version.");
                return;
            }

            _settings.IsSetup = true;
            _settings.Email = Email.Text;
            _settings.ExchangeVersion = (ExchangeVersion)Exchange.SelectedItem;
            _settings.Save();

            LoadSettings();
        }

        /// <summary>
        /// Handles show settings
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ShowSettings_Click(object sender, RoutedEventArgs e)
        {
            if (!_settings.IsSetup)
            {
                return;
            }

            // Default fields to current settings
            Email.Text = _settings.Email;
            Exchange.SelectedItem = _settings.ExchangeVersion;

            // Unmark as setup and reinitialize app
            _settings.IsSetup = false;
            _settings.Save();

            LoadSettings();
        }

        /// <summary>
        /// Handles unpinning/pinning of the window (always on top toggle)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Pin_Click(object sender, RoutedEventArgs e)
        {
            Window parent = Window.GetWindow(this);
            parent.Topmost = !parent.Topmost;
            if (parent.Topmost)
            {
                pin.Content = "unpin";
            }
            else
            {
                pin.Content = "pin";
            }
        }

        #endregion
    }
}