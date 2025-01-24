using Modbus.Device;
using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MKP_Modbus_Data
{
    internal class ModbusRTU
    {
        public ModbusRTU(PortSettings settings) 
        {
            _port = new SerialPort();

            _port.PortName = settings.port_name;
            _port.BaudRate = settings.baud_rate;
            _port.DataBits = settings.data_bits;
            _port.Parity = settings.parity;
            _port.StopBits = settings.stop_bits;
            _port.ReadTimeout = settings.read_timeout;
        }
        public struct PortSettings
        {
            public string port_name;
            public int baud_rate;
            public Parity parity;
            public StopBits stop_bits;
            public int data_bits;
            public int read_timeout;
        }
        public string[] get_com_port_names()
        {
            return System.IO.Ports.SerialPort.GetPortNames();
        }

        public bool connect_modbus_port()
        {
            _port.Open();

            return _port.IsOpen;
        }

        public void create_modbus_master()
        {
            _modbus_master = ModbusSerialMaster.CreateRtu(_port);
        }
        public async Task<ushort[]> read_input_registers(byte slaveId, ushort startAddress, ushort numOfPoints)
        {       
            return await _modbus_master.ReadInputRegistersAsync(slaveId, startAddress, numOfPoints); 
        }

        public async Task<ushort[]> read_holding_registers(byte slaveId, ushort startAddress, ushort numOfPoints)
        {
            return await _modbus_master.ReadHoldingRegistersAsync(slaveId, startAddress, numOfPoints);
        }

        public void disconnect_modbus_port()
        {
            if (_port.IsOpen)
                _port.Close();
        }

        SerialPort _port;
        IModbusSerialMaster _modbus_master;
    }
}
