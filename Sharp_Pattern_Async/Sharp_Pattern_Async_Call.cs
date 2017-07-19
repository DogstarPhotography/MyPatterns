using System;
using System.ComponentModel;

namespace Sharp_Pattern_Async
{
    public class Sharp_Pattern_Async_Call
    {
        public void DoWork()
        { 
            //This is a blocking method
            //Do some work here
        }

        public void StartWork()
        { 
            //This starts the background worker and returns immediately
            MyBackgroundWorker = new BackgroundWorker();
            MyBackgroundWorker.DoWork += new DoWorkEventHandler(MyBackgroundWorker_DoWork);
            MyBackgroundWorker.RunWorkerAsync();
        }
        
        #region Background Worker

        //BackgroundWorker will call DoWork on another thread
        private BackgroundWorker MyBackgroundWorker;

        private void MyBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            //This method calls the blocking DoWork method on another thread thus the StartWork call will not block
            DoWork();
        }

        #endregion
    }
}
