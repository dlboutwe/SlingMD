using System;
using System.Threading.Tasks;
using SlingMD.Outlook.Forms;

namespace SlingMD.Outlook.Services
{
    public class StatusService : IDisposable
    {
        private ProgressForm _progressForm;
        private bool _isDisposed;

        public StatusService()
        {
            _progressForm = new ProgressForm();
            _progressForm.Show();
        }

        public void UpdateProgress(string message, int percentage)
        {
            EnsureNotDisposed();
            _progressForm?.UpdateProgress(message, percentage);
        }

        public void ShowSuccess(string message, bool autoClose = true)
        {
            EnsureNotDisposed();
            _progressForm?.ShowSuccess(message, autoClose);
        }

        public void ShowError(string message, bool autoClose = false)
        {
            EnsureNotDisposed();
            _progressForm?.ShowError(message, autoClose);
        }

        public async Task ShowTemporaryStatusAsync(string message, int durationMs = 3000)
        {
            EnsureNotDisposed();
            UpdateProgress(message, 100);
            await Task.Delay(durationMs);
            _progressForm?.Close();
        }

        private void EnsureNotDisposed()
        {
            if (_isDisposed)
            {
                throw new ObjectDisposedException(nameof(StatusService));
            }
        }

        public void Dispose()
        {
            if (!_isDisposed)
            {
                _progressForm?.Dispose();
                _progressForm = null;
                _isDisposed = true;
            }
        }
    }
} 