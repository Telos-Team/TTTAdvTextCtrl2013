using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Dynamics.Framework.UI.Extensibility;
using Microsoft.Dynamics.Framework.UI.Extensibility.WinForms;

namespace TTTAdvTextCtrl
{
    // Public key token: 1a86c0d41d24fd41

    [ControlAddInExport("TTTAdvTextCtrl")]
    public class AdvTextCtrl : StringControlAddInBase
    {
        private TextBox _textBox;
        private Timer _timer = null;
        private int _maxTimerCount = 0;
        private int _timerCount = 0;

        [ApplicationVisible]
        public event MethodInvoker ControlAddInReady;

        [ApplicationVisible]
        public event MethodInvoker TimerTick;

        [ApplicationVisible]
        public Color TextBoxBackColor
        {
            get { return _textBox.BackColor; }
            set { _textBox.BackColor = value; }
        }

        [ApplicationVisible]
        public Color TextBoxForeColor
        {
            get { return _textBox.ForeColor; }
            set { _textBox.ForeColor = value; }
        }

        [ApplicationVisible]
        public bool VisibleControl
        {
            get { return _textBox.Visible; }
            set { _textBox.Visible = value; }
        }

        [ApplicationVisible]
        public bool VisibleBorder
        {
            get { return _textBox.BorderStyle != BorderStyle.None; }
            set { _textBox.BorderStyle = value ? BorderStyle.FixedSingle : BorderStyle.None; }
        }

        [ApplicationVisible]
        public void ActivateTextBox()
        {
            _textBox.Focus();
        }

        [ApplicationVisible]
        public string OverwriteTextBoxText
        {
            get { return _textBox.Text; }
            set { _textBox.Text = value; }
        }

        [ApplicationVisible]
        public int BoxWidth
        {
            get { return _textBox.Width; }
            set
            {
                _textBox.Anchor = AnchorStyles.Left;
                _textBox.Width = value;
            }
        }

        [ApplicationVisible]
        public Font TextBoxFont
        {
            get { return _textBox.Font; }
            set { _textBox.Font = value; }
        }

        [ApplicationVisible]
        public bool SetProperty(string propertyName, object propertyValue)
        {
            System.Reflection.PropertyInfo pi = _textBox.GetType().GetProperty(propertyName);
            if (!pi.CanWrite) { return false; }
            pi.SetValue(_textBox, propertyValue);

            return true;
        }

        [ApplicationVisible]
        public object GetProperty(string propertyName)
        {
            System.Reflection.PropertyInfo piTextBox = _textBox.GetType().GetProperty(propertyName);
            if (!piTextBox.CanRead) { return null; }
            return piTextBox.GetValue(_textBox);
        }

        [ApplicationVisible]
        public void Update()
        {
            _textBox.Update() ;
        }

        [ApplicationVisible]
        public void StartTimer(int milliseconds)
        {
            if (milliseconds < 1) return;
            StopTimer();

            _timer = new Timer();
            _timer.Tick += Timer_Tick;
            _timer.Interval = milliseconds;
            _timer.Start();
        }
        [ApplicationVisible]
        public void StartTimer(int milliseconds, int maxCount)
        {
            if (maxCount < 0) maxCount = 0;
            StartTimer(milliseconds);
        }

        [ApplicationVisible]
        public void StopTimer()
        {
            if (_timer != null)
            {
                _timer.Stop();
                _timer.Dispose();
                _timer = null;
            }
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            _timer.Stop();

            // Fire event
            if (TimerTick != null)
            {
                TimerTick();
            }

            // Check for max no. of events
            if (_maxTimerCount > 0)
            {
                _timerCount++;
                if (_timerCount == _maxTimerCount)
                {
                    StopTimer();
                    return;
                }
            }

            _timer.Start();
        }


        protected override Control CreateControl()
        {
            _textBox = new TextBox();
            _textBox.HandleCreated += (sender, args) =>
            {
                if (ControlAddInReady != null)
                {
                    ControlAddInReady();
                }
            };
            return _textBox;
        }
    }
}
