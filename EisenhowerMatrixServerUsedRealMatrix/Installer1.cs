using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using ASquare.WindowsTaskScheduler.Models;
using ASquare.WindowsTaskScheduler;
using System.IO.Compression;
using Microsoft.Win32.TaskScheduler;

namespace EisenhowerMatrixServer
{
    [RunInstaller(true)]
    public partial class Installer1 : System.Configuration.Install.Installer
    {
        public Installer1()
       : base()
        {
            this.Committed += new InstallEventHandler(MyInstaller_Committed);
        }

        // Event handler for 'Committed' event.
        private void MyInstaller_Committed(object sender, InstallEventArgs e)
        {
            try
            {
                Directory.SetCurrentDirectory(Path.GetDirectoryName
                (Assembly.GetExecutingAssembly().Location));
                Process.Start(Path.GetDirectoryName(
                  Assembly.GetExecutingAssembly().Location) + "\\EisenhowerMatrixServer.exe");                                              
            }
            catch
            {

            }
        }

        public override void Install(IDictionary stateSaver)
        {
            // Основная установка
            base.Install(stateSaver);

            string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\EisenhowerMatrixServer.exe"; ;
            //Создание задачи в Планировщике задач
            using (var taskService = new TaskService())
            {
                var task = taskService.NewTask();

                task.RegistrationInfo.Description = "Запускает считывание Excel файла каждые 10 минут";

                var trigger = task.Triggers.AddNew(TaskTriggerType.Time);
                trigger.Repetition.Interval = TimeSpan.FromMinutes(10);

                var action = task.Actions.AddNew(TaskActionType.Execute) as ExecAction;
                action.Path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\EisenhowerMatrixServer.exe";

                taskService.RootFolder.RegisterTaskDefinition("Matrix", task);
            }
        }
        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);
            using (var taskService = new TaskService()) taskService.RootFolder.DeleteTask("Matrix");
        }

        public override void Commit(IDictionary savedState)
        {
            base.Commit(savedState);
        }

        // Override the 'Rollback' method.
        public override void Rollback(IDictionary savedState)
        {
            base.Rollback(savedState);
        }
    }
}
