using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace BtMoveAllObjects
{
    class Program
    {
        /// <summary>
        /// BarTenderが終了する前にアプリが落ちたときでもBarTenderが終了するように準備
        /// </summary>
        [DllImport("Kernel32")]
        static extern bool SetConsoleCtrlHandler(HandlerRoutine Handler, bool Add);
        delegate bool HandlerRoutine(CtrlTypes CtrlType);

        private ControlBarTender _ctrlBT;


        private enum CtrlTypes
        {
            CTRL_C_EVENT = 0,
            CTRL_BREAK_EVENT = 1,
            CTRL_CLOSE_EVENT = 2,
            CTRL_LOGOFF_EVENT = 5,
            CTRL_SHUTDOWN_EVENT = 6
        }


        // ハンドラ・ルーチン
        bool myHandler(CtrlTypes ctrlType)
        {
            this._ctrlBT.Dispose();
            return false;
        }


        Program()
        {
            SetConsoleCtrlHandler(myHandler, true);
            this._ctrlBT = new ControlBarTender();
        }


        void Run(string[] args)
        {
            try
            {
                var btInstanceTask = this._ctrlBT.StartBartenderAsync();
                var readConf = ReadConfJson.ReadAsync();

                Task.WhenAll(btInstanceTask, readConf).Wait();

                this._ctrlBT.SetPositionProperties(readConf.Result);

                foreach (string btwfilepath in args)
                {
                    Console.WriteLine("Begin : {0}", btwfilepath);

                    try
                    {
                        this._ctrlBT.ControlBtwFile(btwfilepath);
                        Console.WriteLine("Success : {0}", btwfilepath);
                    }
                    catch (COMException come)
                    {
                        Console.WriteLine("Fail : {0}\n{1}", btwfilepath, come.Message);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Fail : {0}\n{1}", btwfilepath, e.Message);
                    }
                
                }

            }
            catch (COMException come)
            {
                Console.WriteLine(come.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                this._ctrlBT.Dispose();
            }

        }


        static void Main(string[] args)
        {
            (new Program()).Run(args);
        }

    }

}
