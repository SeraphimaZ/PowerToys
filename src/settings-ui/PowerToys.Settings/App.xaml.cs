﻿// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Windows;
using Microsoft.PowerToys.Settings.UI.Library.Telemetry.Events;
using Microsoft.PowerToys.Telemetry;

namespace PowerToys.Settings
{
    /// <summary>
    /// Interaction logic for App.xaml.
    /// </summary>
    public partial class App : Application
    {
        private MainWindow settingsWindow;

        public bool ShowOobe { get; set; }

        public void OpenSettingsWindow(Type type)
        {
            if (settingsWindow == null)
            {
                settingsWindow = new MainWindow();
            }

            settingsWindow.Open(type);
        }

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            if (!ShowOobe)
            {
                settingsWindow = new MainWindow();
                settingsWindow.Show();
            }
            else
            {
                PowerToysTelemetry.Log.WriteEvent(new OobeStartedEvent());

                OobeWindow otherWindow = new OobeWindow();
                otherWindow.Show();
            }
        }
    }
}
