using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Word = Microsoft.Office.Interop.Word;

namespace OfficeDemo
{

    class Report : IDisposable
    {
        string outputPath;
        private bool disposed = false;

        public Report(string templatePath, string outputPath, Word._Application app) 
        {
            wordApp = app;
            this.outputPath = outputPath;
            loadDocFromPath(templatePath);
        }
        private Word._Application wordApp { set; get; }

        private Word._Document document { set; get; }

        private void loadDocFromPath(string path) {
            document = wordApp.Documents.Open(path);
        }

        public void SetBookmarkNamed(string bookmark, string text) {
            var place = document.Bookmarks[bookmark];
            place.Range.Text = text;
        }

        public void SaveOut() {
            checkIfDisposed();
            document.SaveAs2(outputPath);
        }

        public void Dispose() {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected void Dispose(bool disposing) {
            if (disposing) {
                document.Close();
                document = null;
            }

            this.disposed = true;
        }

        ~Report() {
            this.Dispose(disposing: false);
        }

        private void checkIfDisposed() {
            if (this.disposed)
            {
                throw new ObjectDisposedException("The report doc has been disposed of.");
            }
        }


    }


}
