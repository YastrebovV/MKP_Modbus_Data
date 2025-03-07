﻿using System;
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
        bool connected = false;

        private async Task<ushort> get_program_state(byte slave_id, ushort state_address)
        {
            ushort[] program_run = new ushort[1];
            try
            {
                program_run = await _modbusRTU.read_input_registers(slave_id, state_address, 1).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                program_run[0] = 0;
            }
           return program_run[0];
        }
        private async Task<ushort[]> get_data_from_MKP(byte slave_id, ushort pv_address, ushort sp_address)
        {
            ushort[] MKP_data = new ushort[2];
            ushort[] pv = null;
            ushort[] sp = null;
            try
            {
                pv = await _modbusRTU.read_input_registers(slave_id, pv_address, 1).ConfigureAwait(false);
                sp = await _modbusRTU.read_input_registers(slave_id, sp_address, 1).ConfigureAwait(false);

                MKP_data[0] = pv[0];
                MKP_data[1] = sp[0]; 
            }
            catch (Exception)
            {
                MKP_data[0] = 0;
                MKP_data[1] = 0;
            }
            return MKP_data;
        }
        private async Task get_data_from_all_mkp(CancellationToken cancel_token)
        {
            while (true)
            {
                if (cancel_token.IsCancellationRequested)
                {
                    return;
                }

                ushort[] MKP_1_data = null;
                ushort[] MKP_2_data = null;
                ushort[] MKP_3_data = null;
                ushort[] MKP_4_data = null;
                ushort[] MKP_5_data = null;

                ushort MKR_1_state = await get_program_state(1, 91);
                ushort MKR_2_state = await get_program_state(2, 91);
                ushort MKR_3_state = await get_program_state(3, 91);
                ushort MKR_4_state = await get_program_state(4, 91);
                ushort MKR_5_state = await get_program_state(5, 91);

                MKP_1_data = await get_data_from_MKP(1, 123, 137);
                MKP_2_data = await get_data_from_MKP(2, 123, 137);
                MKP_3_data = await get_data_from_MKP(3, 123, 137);
                MKP_4_data = await get_data_from_MKP(4, 123, 137);
                MKP_5_data = await get_data_from_MKP(5, 123, 137);

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
                    await _plc_control.write_dbw_async(_ini.IniReadValue("PLC_Info", "Mkp1RunState"), MKR_1_state);
                    await _plc_control.write_dbw_async(_ini.IniReadValue("PLC_Info", "Mkp1PV"), MKP_1_data[0]);
                    await _plc_control.write_dbw_async(_ini.IniReadValue("PLC_Info", "Mkp1SP"), MKP_1_data[1]);

                    await _plc_control.write_dbw_async(_ini.IniReadValue("PLC_Info", "Mkp2RunState"), MKR_2_state);
                    await _plc_control.write_dbw_async(_ini.IniReadValue("PLC_Info", "Mkp2PV"), MKP_2_data[0]);
                    await _plc_control.write_dbw_async(_ini.IniReadValue("PLC_Info", "Mkp2SP"), MKP_2_data[1]);

                    await _plc_control.write_dbw_async(_ini.IniReadValue("PLC_Info", "Mkp3RunState"), MKR_3_state);
                    await _plc_control.write_dbw_async(_ini.IniReadValue("PLC_Info", "Mkp3PV"), MKP_3_data[0]);
                    await _plc_control.write_dbw_async(_ini.IniReadValue("PLC_Info", "Mkp3SP"), MKP_3_data[1]);

                    await _plc_control.write_dbw_async(_ini.IniReadValue("PLC_Info", "Mkp4RunState"), MKR_4_state);
                    await _plc_control.write_dbw_async(_ini.IniReadValue("PLC_Info", "Mkp4PV"), MKP_4_data[0]);
                    await _plc_control.write_dbw_async(_ini.IniReadValue("PLC_Info", "Mkp4SP"), MKP_4_data[1]);

                    await _plc_control.write_dbw_async(_ini.IniReadValue("PLC_Info", "Mkp5RunState"), MKR_5_state);
                    await _plc_control.write_dbw_async(_ini.IniReadValue("PLC_Info", "Mkp5PV"), MKP_5_data[0]);
                    await _plc_control.write_dbw_async(_ini.IniReadValue("PLC_Info", "Mkp5SP"), MKP_5_data[1]);                    
                }

                Thread.Sleep(1000);
            }
        }
        private void Main_Load(object sender, EventArgs e)
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
        }
        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            _plc_control.disconnect();
            _modbusRTU.disconnect_modbus_port();
        }

        private void but_port_connect_click(object sender, EventArgs e)
        {
           
            if (!connected)
            {
                if (_modbusRTU.connect_modbus_port())
                {
                    _modbusRTU.create_modbus_master();
                    but_port_connect.Text = "Отключиться от Com порта";
                    connected = true;
                }
            }
            else
            {
                _modbusRTU.disconnect_modbus_port();
                but_port_connect.Text = "Подключиться к Com порту";
                connected = false;
            }
        }

        private async void but_turn_on_data_coll_Click(object sender, EventArgs e)
        {
            if (tasks_started)
            {
                _cancel_token.Cancel();
                tasks_started = false;
                but_turn_on_data_coll.Text = "Включить сбор \n данных";
                _plc_control.disconnect();

                return;
            }

            if (connected && !tasks_started)
            {
                _cancel_token = new CancellationTokenSource();
                tasks_started = true;
                but_turn_on_data_coll.Text = "Отключить сбор \n данных";
                
                await _plc_control.connect_async(_cancel_token.Token).ConfigureAwait(continueOnCapturedContext: false); ;
                await get_data_from_all_mkp(_cancel_token.Token).ConfigureAwait(continueOnCapturedContext: false); ;
                return;
            }
        }

        private void close_but_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void min_but_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
    }
}
