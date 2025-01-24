using S7.Net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MKP_Modbus_Data
{
    public class PLC_S7
    {
        public PLC_S7(CpuType my_cpu_type, String my_cpu_ip, short my_cpu_rack, short my_cpu_slot)
        {
            _my_cpu_type = my_cpu_type;
            _my_cpu_ip = my_cpu_ip;
            _my_cpu_rack = my_cpu_rack;
            _my_cpu_slot = my_cpu_slot;
        }
        public async Task connect_async(CancellationToken cancel_token)
        {
            _plc = new Plc(_my_cpu_type, _my_cpu_ip, _my_cpu_rack, _my_cpu_slot);
           
            try
            {        
                await _plc.OpenAsync(cancel_token).ConfigureAwait(continueOnCapturedContext: false); //попытка подключения к plc
            }
            catch (Exception plc_ex)
            {
              //  MessageBox.Show(plc_ex.Message, "Ошибка подключения", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
           
        }  //функция подключения к плс
        public void disconnect()
        {
            if (_plc == null)
                return;

            if (_plc.IsConnected)
                _plc.Close();
        }

        public async Task<object> read_dbw_async(string variable)
        {
            return await _plc.ReadAsync(variable);
        }
        public async Task write_dbw_async(string variable, UInt16 new_value)
        {
            await _plc.WriteAsync(variable, new_value);
        }

        public bool get_state_plc_connected()
        {
            return _plc.IsConnected;
        }

        CpuType _my_cpu_type;
        String _my_cpu_ip; 
        short _my_cpu_rack;
        short _my_cpu_slot;
        Plc _plc;
    }
}
