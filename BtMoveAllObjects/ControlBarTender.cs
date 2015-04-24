using BarTender;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Diagnostics;

namespace BtMoveAllObjects
{

    class ControlBarTender : IDisposable
    {
        private Application _btApp;
        private BtUnits _btUnits;
        private double _x;
        private double _y;
        private int _btAppProcessId;

        public ControlBarTender()
        {
            this._btUnits = BtUnits.btUnitsCurrent;
            this._x = 0.0;
            this._y = 0.0;
            this._btAppProcessId = -1;
        }

        ~ControlBarTender()
        {
            Dispose(false);
        }


        public async Task StartBartenderAsync()
        {
            await Task.Run(() =>
            {
                this._btApp = new BarTender.Application()
                {
                    VisibleWindows = BarTender.BtVisibleWindows.btNone
                };

                this._btAppProcessId = this._btApp.ProcessId;
            });
        }

        public void StartBartender()
        {
            this._btApp = new Application()
            {
                VisibleWindows = BarTender.BtVisibleWindows.btNone
            };

            this._btAppProcessId = this._btApp.ProcessId;
        }

        public void SetPositionProperties(JsonConfigs conf)
        {
            switch (conf.Unit)
            {
                case "mm":
                    this._btUnits = BarTender.BtUnits.btUnitsMillimeters;
                    break;

                case "cm":
                    this._btUnits = BarTender.BtUnits.btUnitsCentimeters;
                    break;

                case "inch":
                    this._btUnits = BarTender.BtUnits.btUnitsInches;
                    break;

                default:
                    this._btUnits = BarTender.BtUnits.btUnitsCurrent;
                    break;
            }

            this._x = conf.X;
            this._y = conf.Y;
        }


        ///<summary>
        ///btwfileをOpen
        ///</summary>
        public void ControlBtwFile(string btwfilepath)
        {
            try
            {
                using (var hProcess = Process.GetProcessById(this._btAppProcessId))
                {
                }
            }
            catch (ArgumentException)
            {
                Console.WriteLine("Restart BarTender");
                Marshal.FinalReleaseComObject(this._btApp);
                StartBartender();
            }

            var btFormat = this._btApp.Formats.Open(btwfilepath);

            btFormat.MeasurementUnits = this._btUnits;

            Enumerable.Range(1, btFormat.Objects.Count)
                .Select(i => btFormat.Objects.Item(i))
                .Where(xs => xs.Type != BarTender.BtObjectType.btObjectError
                    && xs.Type != BarTender.BtObjectType.btObjectRFID
                    && xs.Type != BarTender.BtObjectType.btObjectMagneticStripe
                    && xs.Type != BarTender.BtObjectType.btObjectShape
                    )
                    .ForEach(xs =>
                    {
                        xs.X += this._x;
                        xs.Y += this._y;
                    }
                    );

            btFormat.Close(BarTender.BtSaveOptions.btSaveChanges);
        }


        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }


        protected virtual void Dispose(bool disposing)
        {
            //for managed resources
            if (disposing)
            { }

            //for unmanaged resources
            if (this._btApp != null)
            {
                this._btApp.Quit(BtSaveOptions.btDoNotSaveChanges);
                Marshal.FinalReleaseComObject(this._btApp);
            }

        }

    }
}
