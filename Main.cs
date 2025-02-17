using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Modbus.Data;
using Modbus.Device;
using Modbus.Utility;

using System.Threading;
using System.Diagnostics.Eventing.Reader;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using System.Runtime.InteropServices;

using S7.Net;
using S7.Net.Types;
using THM_Arc;

namespace MKP_Modbus_Data
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        ModbusRTU _modbusRTU;

        CancellationTokenSource _cancel_token;
        IModbusSerialMaster _modbus_master;

        PLC_S7 _plc_control;

        IniClass _ini = new IniClass();

        bool tasks_started = false;  
        //bool connected = false;

        string mkp1_run_status_db_adr;
        string mkp1_pv_db_adr;
        string mkp1_sp_db_adr;
        string mkp2_run_status_db_adr;
        string mkp2_pv_db_adr;
        string mkp2_sp_db_adr;
        string mkp3_run_status_db_adr;
        string mkp3_pv_db_adr;
        string mkp3_sp_db_adr;
        string mkp4_run_status_db_adr;
        string mkp4_pv_db_adr;
        string mkp4_sp_db_adr;
        string mkp5_run_status_db_adr;
        string mkp5_pv_db_adr;
        string mkp5_sp_db_adr;

        string life_bit_db_adr;

        /*
         * Загрузка данных из ini файла
         */
        private void load_data_from_inifile()
        {
            mkp1_run_status_db_adr = _ini.IniReadValue("PLC_Info", "Mkp1RunState");
            mkp1_pv_db_adr = _ini.IniReadValue("PLC_Info", "Mkp1PV");
            mkp1_sp_db_adr = _ini.IniReadValue("PLC_Info", "Mkp1SP");

            mkp2_run_status_db_adr = _ini.IniReadValue("PLC_Info", "Mkp2RunState");
            mkp2_pv_db_adr = _ini.IniReadValue("PLC_Info", "Mkp2PV");
            mkp2_sp_db_adr = _ini.IniReadValue("PLC_Info", "Mkp2SP");

            mkp3_run_status_db_adr = _ini.IniReadValue("PLC_Info", "Mkp3RunState");
            mkp3_pv_db_adr = _ini.IniReadValue("PLC_Info", "Mkp3PV");
            mkp3_sp_db_adr = _ini.IniReadValue("PLC_Info", "Mkp3SP");

            mkp4_run_status_db_adr = _ini.IniReadValue("PLC_Info", "Mkp4RunState");
            mkp4_pv_db_adr = _ini.IniReadValue("PLC_Info", "Mkp4PV");
            mkp4_sp_db_adr = _ini.IniReadValue("PLC_Info", "Mkp4SP");

            mkp5_run_status_db_adr = _ini.IniReadValue("PLC_Info", "Mkp5RunState");
            mkp5_pv_db_adr = _ini.IniReadValue("PLC_Info", "Mkp5PV");
            mkp5_sp_db_adr = _ini.IniReadValue("PLC_Info", "Mkp5SP");

            life_bit_db_adr = _ini.IniReadValue("PLC_Info", "LifeBit");
        }
        /*
         * Получение данных с терморегулятора
         */
        private ushort[] get_data_from_mkp(byte slave_id, ushort pv_address, ushort sp_address, ushort state_address)
        {
            ushort[] MKP_data = new ushort[3];
            ushort[] pv = null;
            ushort[] sp = null;
            ushort[] state= null;
            try
            {
                pv =  _modbusRTU.read_input_registers(slave_id, pv_address, 1);
                sp =  _modbusRTU.read_input_registers(slave_id, sp_address, 1);
                state = _modbusRTU.read_input_registers(slave_id, state_address, 1);

                MKP_data[0] = pv[0];
                MKP_data[1] = sp[0];
                MKP_data[2] = state[0];
            }
            catch (Exception)
            {
                MKP_data[0] = 0;
                MKP_data[1] = 0;
                MKP_data[2] = 0;
            }
            return MKP_data;
        }
        /*
         * Получение и обработка данных с терморегуляторов
         */
        private async Task get_data_from_all_mkp()
        {
            await get_data_from_all_mkp(_cancel_token.Token).ConfigureAwait(continueOnCapturedContext: false);
        }
        private async Task get_data_from_all_mkp(CancellationToken cancel_token)
        {
            ushort lenght_arr = 3;
            while (true)
            {
                if (cancel_token.IsCancellationRequested)
                {
                    return;
                }

                ushort[] MKP_1_data = new ushort[lenght_arr];
                ushort[] MKP_2_data = new ushort[lenght_arr];
                ushort[] MKP_3_data = new ushort[lenght_arr];
                ushort[] MKP_4_data = new ushort[lenght_arr];
                ushort[] MKP_5_data = new ushort[lenght_arr];

                MKP_1_data =  get_data_from_mkp(1, 123, 137, 91);
                MKP_2_data =  get_data_from_mkp(2, 123, 137, 91);
                MKP_3_data =  get_data_from_mkp(3, 123, 137, 91);
                MKP_4_data =  get_data_from_mkp(4, 123, 137, 91);
                MKP_5_data =  get_data_from_mkp(5, 123, 137, 91);

                lb_pv_id1.Invoke((Action)delegate { lb_pv_id1.Text = MKP_1_data[0].ToString(); });
                lb_sp_id1.Invoke((Action)delegate { lb_sp_id1.Text = MKP_1_data[1].ToString(); });
      
                lb_pv_id2.Invoke((Action)delegate { lb_pv_id2.Text = MKP_2_data[0].ToString(); });
                lb_sp_id2.Invoke((Action)delegate { lb_sp_id2.Text = MKP_2_data[1].ToString(); });
                
                lb_pv_id3.Invoke((Action)delegate { lb_pv_id3.Text = MKP_3_data[0].ToString(); });
                lb_sp_id3.Invoke((Action)delegate { lb_sp_id3.Text = MKP_3_data[1].ToString(); });              

                lb_pv_id4.Invoke((Action)delegate { lb_pv_id4.Text = MKP_4_data[0].ToString(); });
                lb_sp_id4.Invoke((Action)delegate { lb_sp_id4.Text = MKP_4_data[1].ToString(); });
           
                lb_pv_id5.Invoke((Action)delegate { lb_pv_id5.Text = MKP_5_data[0].ToString(); });
                lb_sp_id5.Invoke((Action)delegate { lb_sp_id5.Text = MKP_5_data[1].ToString(); });

                if (_plc_control.get_state_plc_connected())
                {
                    await _plc_control.write_db_async(mkp1_run_status_db_adr, MKP_1_data[2]);
                    await _plc_control.write_db_async(mkp1_pv_db_adr, MKP_1_data[0]);
                    await _plc_control.write_db_async(mkp1_sp_db_adr, MKP_1_data[1]);

                    await _plc_control.write_db_async(mkp2_run_status_db_adr, MKP_2_data[2]);
                    await _plc_control.write_db_async(mkp2_pv_db_adr, MKP_2_data[0]);
                    await _plc_control.write_db_async(mkp2_sp_db_adr, MKP_2_data[1]);

                    await _plc_control.write_db_async(mkp3_run_status_db_adr, MKP_3_data[2]);
                    await _plc_control.write_db_async(mkp3_pv_db_adr, MKP_3_data[0]);
                    await _plc_control.write_db_async(mkp3_sp_db_adr, MKP_3_data[1]);

                    await _plc_control.write_db_async(mkp4_run_status_db_adr, MKP_4_data[2]);
                    await _plc_control.write_db_async(mkp4_pv_db_adr, MKP_4_data[0]);
                    await _plc_control.write_db_async(mkp4_sp_db_adr, MKP_4_data[1]);

                    await _plc_control.write_db_async(mkp5_run_status_db_adr, MKP_5_data[2]);
                    await _plc_control.write_db_async(mkp5_pv_db_adr, MKP_5_data[0]);
                    await _plc_control.write_db_async(mkp5_sp_db_adr, MKP_5_data[1]);


                    if (Convert.ToBoolean(await _plc_control.read_db_async(life_bit_db_adr)))
                    {
                        await _plc_control.write_db_async(life_bit_db_adr, false);
                    }
                    else
                    {
                        await _plc_control.write_db_async(life_bit_db_adr, true);
                    }
                }
                Task.Delay(1000).Wait();
            }
        }
        /*
         * Метод обработки загрузки формы
         */

        private async Task connect_async_to_plc()
        {
            await _plc_control.connect_async(_cancel_token.Token).ConfigureAwait(continueOnCapturedContext: false);
        }
        private void main_load(object sender, EventArgs e)
        {
            CpuType _cpu_Type = CpuType.S7300;
            ModbusRTU.PortSettings _port_settings = new ModbusRTU.PortSettings();

            string path = System.Windows.Forms.Application.StartupPath + "\\IniFile.ini";
            _ini.IniFile(path);

            int type_PLC = Convert.ToInt32(_ini.IniReadValue("PLC_Info", "typePLC"));

            switch (type_PLC)
            {
                case 200: _cpu_Type = CpuType.S7200; break;
                case 300: _cpu_Type = CpuType.S7300; break;
                case 400: _cpu_Type = CpuType.S7400; break;
                case 1200: _cpu_Type = CpuType.S71200; break;
                case 1500: _cpu_Type = CpuType.S71500; break;
            }

            /*Загружаем параметры Com порта с ини файла*/
            _port_settings.port_name = _ini.IniReadValue("Port_Info", "PortName");
            _port_settings.baud_rate = Convert.ToInt32(_ini.IniReadValue("Port_Info", "BaudRate"));

            switch (_ini.IniReadValue("Port_Info", "Parity"))
            {
                case "None":
                    _port_settings.parity = Parity.None; break;
                case "Odd":
                    _port_settings.parity = Parity.Odd; break;
                case "Mark":
                    _port_settings.parity = Parity.Mark; break;
                case "Even":
                    _port_settings.parity = Parity.Even; break;
                case "Space":
                    _port_settings.parity = Parity.Space; break;
            }

            switch (_ini.IniReadValue("Port_Info", "StopBits"))
            {
                case "One":
                    _port_settings.stop_bits = StopBits.One; break;
                case "Two":
                    _port_settings.stop_bits = StopBits.Two; break;
                case "OnePointFive":
                    _port_settings.stop_bits = StopBits.OnePointFive; break;
                case "None":
                    _port_settings.stop_bits = StopBits.None; break;
            }

            _port_settings.data_bits = Convert.ToInt16(_ini.IniReadValue("Port_Info", "DataBits"));
            _port_settings.read_timeout = Convert.ToInt32(_ini.IniReadValue("Port_Info", "ReadTimeout"));
            /********/

            _modbusRTU = new ModbusRTU(_port_settings);

            _plc_control = new PLC_S7(_cpu_Type, _ini.IniReadValue("PLC_Info", "IP"),
                Convert.ToInt16(_ini.IniReadValue("PLC_Info", "rack")), Convert.ToInt16(_ini.IniReadValue("PLC_Info", "slot")));

            _cancel_token = new CancellationTokenSource();

            load_data_from_inifile();
          
            try
            {
                /*Попытка подключения к com порту*/
                but_port_connect.PerformClick();

                if (_modbusRTU.get_port_state())
                {              
                    /*
                     * Если подключение к com порту прошло успешно, 
                     * запускаем метод сбора данных с датчиков
                     */
                    but_turn_on_data_coll.PerformClick();
                    /*
                    * Попытка подключения к ПЛК
                    */
                    but_connect_plc.PerformClick();
                }              
            }
            catch (Exception){}
         
        }
        /*
         * Метод обработки закрытия формы
         */
        private void main_form_closing(object sender, FormClosingEventArgs e)
        {
            _plc_control.disconnect();
            _modbusRTU.disconnect_modbus_port();
        }
        private void but_port_connect_click(object sender, EventArgs e)
        {
            if (!_modbusRTU.get_port_state())
            {
                if (_modbusRTU.connect_modbus_port())
                {
                    _modbusRTU.create_modbus_master();
                    but_port_connect.Text = "Отключиться от Com порта";
                    timer_check_connect_com.Start();
                }
            }
            else
            {
                _modbusRTU.disconnect_modbus_port();
                but_port_connect.Text = "Подключиться к Com порту";
                timer_check_connect_com.Stop();
            }
        }
        private async void but_connect_plc_click(object sender, EventArgs e)
        {
            if (!_plc_control.get_state_plc_connected())
            {
                await Task.Run(connect_async_to_plc);
                but_connect_plc.Text = "Отключиться \r\nот ПЛК";
                timer_check_connect_to_plc.Start();
            }
            else
            {
                _plc_control.disconnect();
                but_connect_plc.Text = "Подключиться \r\nк ПЛК";
                timer_check_connect_to_plc.Stop();
            }         
        }
        private async void but_turn_on_data_coll_click(object sender, EventArgs e)
        {
            if (tasks_started)
            {
                _cancel_token.Cancel();
                tasks_started = false;
                but_turn_on_data_coll.Text = "Включить сбор \n данных";
                _plc_control.disconnect();

                return;
            }

            if (_modbusRTU.get_port_state() && !tasks_started)
            {
                _cancel_token = new CancellationTokenSource();
                tasks_started = true;
                but_turn_on_data_coll.Text = "Отключить сбор \n данных";

                await Task.Run(get_data_from_all_mkp);          
                await Task.Run(connect_async_to_plc);
                return;
            }
        }
        private void close_but_click(object sender, EventArgs e)
        {
            Close();
        }
        private void min_but_click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        private void main_mouse_down(object sender, MouseEventArgs e)
        {
            base.Capture = false;
            Message m = Message.Create(base.Handle, 0xa1, new IntPtr(2), IntPtr.Zero);
            this.WndProc(ref m);
        }
        private void timer_check_connect_to_plc_Tick(object sender, EventArgs e)
        {
            if (!_plc_control.get_state_plc_connected())
            {
                timer_check_connect_to_plc.Stop();
                but_connect_plc.Text = "Подключиться \r\nк ПЛК";
            }
        }
        private void timer_check_connect_com_Tick(object sender, EventArgs e)
        {
            if (!_modbusRTU.get_port_state())
            {
                timer_check_connect_com.Stop();
                but_port_connect.Text = "Подключиться к Com порту";       
            }
        }
    }
}
