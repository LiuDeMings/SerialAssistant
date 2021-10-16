using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;

namespace SerialAssistant
{
    public partial class Form1 : Form
    {
        private long receive_count = 0; //接收字节计数
        private long send_count = 0;    //发送字节计数
        private StringBuilder sb = new StringBuilder();    //为了避免在接收处理函数中反复调用，依然声明为一个全局变量


        public Form1()
        {
            InitializeComponent();
            timer3.Interval = 10;
            timer3.Start();
        }

        private bool search_port_is_exist(String item, String[] port_list)
        {
            for (int i = 0; i < port_list.Length; i++)
            {
                if (port_list[i].Equals(item))
                {
                    return true;
                }
            }

            return false;
        }

        /* 扫描串口列表并添加到选择框 */
        private void Update_Serial_List()
        {
            try
            {
                /* 搜索串口 */
                String[] cur_port_list = System.IO.Ports.SerialPort.GetPortNames();

                /* 刷新串口列表comboBox */
                int count = comboBox1.Items.Count;
                if (count == 0)
                {
                    //combox中无内容，将当前串口列表全部加入
                    comboBox1.Items.AddRange(cur_port_list);
                    return;
                }
                else
                {
                    //判断有无新插入的串口
                    for (int i = 0; i < cur_port_list.Length; i++)
                    {
                        if (!comboBox1.Items.Contains(cur_port_list[i]))
                        {
                            //找到新插入串口，添加到combox中
                            comboBox1.Items.Add(cur_port_list[i]);
                        }
                    }

                    //判断有无拔掉的串口
                    for (int i = 0; i < count; i++)
                    {
                        if (!search_port_is_exist(comboBox1.Items[i].ToString(), cur_port_list))
                        {
                            //找到已被拔掉的串口，从combox中移除
                            comboBox1.Items.RemoveAt(i);
                        }
                    }
                }

                /* 如果当前选中项为空，则默认选择第一项 */
                if (comboBox1.Items.Count > 0)
                {
                    if (comboBox1.Text.Equals(""))
                    {
                        //软件刚启动时，列表项的文本值为空
                        comboBox1.Text = comboBox1.Items[0].ToString();
                    }
                }
                else
                {
                    //无可用列表，清空文本值
                    comboBox1.Text = "";
                }
            }
            catch (Exception)
            {
                //当下拉框被打开时，修改下拉框会发生异常
                return;
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            /* 添加串口选择列表 */
            Update_Serial_List();

            /* 添加波特率列表 */
            string[] baud = { "9600", "38400", "57600", "115200" };
            comboBox2.Items.AddRange(baud);

            /* 添加数据位列表 */
            string[] data_length = { "5", "6", "7", "8", "9" };
            comboBox3.Items.AddRange(data_length);

            /* 添加校验位列表 */
            string[] verification_mode = { "None", "Odd", "Even", "Mark", "Space" };
            comboBox4.Items.AddRange(verification_mode);

            /* 添加停止位列表 */
            string[] stop_length = { "1", "1.5", "2" };
            comboBox5.Items.AddRange(stop_length);


            /* 在串口未打开的情况下每隔500ms刷新一次串口列表框 */
            timer1.Interval = 500;
            timer1.Start();
        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            Update_Serial_List();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //将可能产生异常的代码放置在try块中
                //根据当前串口属性来判断是否打开
                if (serialPort1.IsOpen)
                {
                    //串口已经处于打开状态

                    serialPort1.Close();    //关闭串口
                    button1.BackgroundImage = global::SerialAssistant.Properties.Resources.connect;
                    comboBox1.Enabled = true;
                    comboBox2.Enabled = true;
                    comboBox3.Enabled = true;
                    comboBox4.Enabled = true;
                    comboBox5.Enabled = true;
                    label6.Text = "串口已关闭!";
                    label6.ForeColor = Color.Red;

                    //开启端口扫描
                    timer1.Interval = 1000;
                    timer1.Start();
                }
                else
                {
                    /* 串口已经处于关闭状态，则设置好串口属性后打开 */
                    //停止串口扫描
                    timer1.Stop();

                    comboBox1.Enabled = false;
                    comboBox2.Enabled = false;
                    comboBox3.Enabled = false;
                    comboBox4.Enabled = false;
                    comboBox5.Enabled = false;

                    serialPort1.PortName = comboBox1.Text;
                    serialPort1.BaudRate = Convert.ToInt32(comboBox2.Text);
                    serialPort1.DataBits = Convert.ToInt16(comboBox3.Text);

                    if (comboBox4.Text.Equals("None"))
                        serialPort1.Parity = System.IO.Ports.Parity.None;
                    else if (comboBox4.Text.Equals("Odd"))
                        serialPort1.Parity = System.IO.Ports.Parity.Odd;
                    else if (comboBox4.Text.Equals("Even"))
                        serialPort1.Parity = System.IO.Ports.Parity.Even;
                    else if (comboBox4.Text.Equals("Mark"))
                        serialPort1.Parity = System.IO.Ports.Parity.Mark;
                    else if (comboBox4.Text.Equals("Space"))
                        serialPort1.Parity = System.IO.Ports.Parity.Space;

                    if (comboBox5.Text.Equals("1"))
                        serialPort1.StopBits = System.IO.Ports.StopBits.One;
                    else if (comboBox5.Text.Equals("1.5"))
                        serialPort1.StopBits = System.IO.Ports.StopBits.OnePointFive;
                    else if (comboBox5.Text.Equals("2"))
                        serialPort1.StopBits = System.IO.Ports.StopBits.Two;

                    //打开串口，设置状态
                    serialPort1.Open();
                    button1.BackgroundImage = global::SerialAssistant.Properties.Resources.disconnect;
                    label6.Text = "串口已打开!";
                    label6.ForeColor = Color.Green;

                }
            }
            catch (Exception ex)
            {
                //捕获可能发生的异常并进行处理

                //捕获到异常，创建一个新的对象，之前的不可以再用  
                serialPort1 = new System.IO.Ports.SerialPort(components);
                serialPort1.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(serialPort1_DataReceived);

                //响铃并显示异常给用户
                System.Media.SystemSounds.Beep.Play();
                button1.BackgroundImage = global::SerialAssistant.Properties.Resources.connect;
                MessageBox.Show(ex.Message);
                comboBox1.Enabled = true;
                comboBox2.Enabled = true;
                comboBox3.Enabled = true;
                comboBox4.Enabled = true;
                comboBox5.Enabled = true;
                label6.Text = "串口已关闭!";
                label6.ForeColor = Color.Red;

                //开启串口扫描
                timer1.Interval = 500;
                timer1.Start();
            }
        }
        List<byte> buffer = new List<byte>(1024);
        private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e) 
        {

            int num = serialPort1.BytesToRead;      //获取接收缓冲区中的字节数
            byte[] buf = new byte[num];             //声明一个大小为num的字节数据用于存放读出的byte型数据


            receive_count += num;                   //接收字节计数变量增加nun
            serialPort1.Read(buf, 0, num);          //读取接收缓冲区中num个字节到byte数组中

            buffer.AddRange(buf);                   //不断地将接收到的数据加入到buffer链表中

            try
            {
                while (buffer.Count >= 4)
                {

                    if (buffer[0] == 0x10 && buffer[1] == 0x10) //传输数据有帧头，用于判断
                    {


                        if (buffer.Count < 7) //数据区尚未接收完整
                        {
                            break;
                        }
                        Invoke((EventHandler)(delegate
                        {
                            //for (UInt16 i = 0; i < buffer.Count; i++)
                            //{
                            //    textBox3.AppendText(i.ToString() + "=" + buffer[i].ToString() + "  ");

                            //}
                            RxDataProcess(buffer.ToArray(), buffer.Count);
                            buffer.RemoveRange(0, 7);
                        }
                     )
                     );
                    }
                    else
                    {
                        buffer.RemoveAt(0);//清除第一个字节，继续检测下一个。
                    }


                }
            }
            catch (Exception ex)
            {
                //响铃并显示异常给用户
                System.Media.SystemSounds.Beep.Play();
                MessageBox.Show(ex.Message);

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            receive_count = 0;          //接收计数清零
            send_count = 0;             //发送计数
            label8.Text = "接收：" + receive_count.ToString() + " Bytes";
            label7.Text = "发送：" + receive_count.ToString() + " Bytes";
        }

        public void RxDataProcess(byte[] buff, int len)
        {
            //textBox3.AppendText("\r\n");
            //for (UInt16 i = 0; i < len; i++)           //打印获取的数据
            //{
            //    textBox3.AppendText(i.ToString("X2") + "=" + buff[i].ToString() + " ");
            //}
            //textBox3.AppendText("\r\n");
            if (len != 8)      //数据长度检查
            {
                textBox3.AppendText("len=" + len.ToString());
                return;
            }

            UInt16 head = (UInt16)(Convert.ToUInt32(MessageHeadBox.Text, 16));  //数据头检查
            if (buff[0] != (Byte)((head >> 8) & 0xff) || buff[1] != (Byte)((head >> 8) & 0xff))
            {
                return;
            }
            UInt16 crc_check = Crc16(buff, 6);;          //CRC16校验码  
            UInt16 rx_crc = (UInt16)(buff[6] + (buff[7] << 8));
            //  textBox3.AppendText("crc_sum="+ crc_check.ToString());
            //  textBox3.AppendText("get crc_sum=" + rx_crc.ToString());
            if (rx_crc != crc_check)
            {
                return;
            }

            rx_index = (UInt16)(buff[2] + (buff[3] << 8));

            if (rx_index == 0)  mcu_ready_ok = 1;                   //下位机完成升级准备
            if (rx_index == pack_count) update_status = 4;          //下位机完成升级
            rx_flag = 1;                                            //接收到返回数据，标志置1
            rx_time = System.DateTime.Now;                          //获取当前时间
            failed_transmit_cnt = 0;
            textBox3.AppendText("\r\n[" + rx_time.ToString("HH:mm:ss:fff") + "]" + ":收到第 " + rx_index.ToString() + " 包数据反馈");
            TimeSpan span = (TimeSpan)(rx_time - tx_time);
            textBox3.AppendText("\r\ndelta=" + span.Milliseconds.ToString());
 

         

        }


        private void button3_Click(object sender, EventArgs e)
        {
            string file;

            /* 弹出文件选择框供用户选择 */
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;//该值确定是否可以选择多个文件
            dialog.Title = "请选择要加载的文件(文本格式)";
            dialog.Filter = "文本文件(*.txt)|*.txt|Bin文件(*.bin)|*.bin";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                file = dialog.FileName;
            }
            else
            {
                return;
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            //清空发送缓冲区
            textBox2.Text = "";
            textBox3.Text = "";
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            //清空发送缓冲区
            textBox2.Text = comboBox7.SelectedItem.ToString();
        }


        private void timer2_Tick(object sender, EventArgs e)
        {
            
        }





        /*                           数据结构
         *   消息头   升级包序号  固件数据长度  CRC校验低位 CRC校验高位
         *   2字节      1字节      X（1~256）字节      1字节        2字节
         * 
         *   整包数据长度=X+6字节
         * 
         */

        //文件
        private static int response_max_time = 900;   //允许响应的最长时间，超过则认为下位机未响应
        private static int failed_transmit_cnt = 0;  //发送失败计数
        private static int update_status = 0;         //0无动作  1发送准备请求  2等待下位机返回准备完成状态 3已接收到完成状态进行升级 4发送完最后一包数据 5升级完成 
        private static FileStream fs = null;   
        private static long tx_index = 0;            //发送数据包序号
        private static long rx_index = 0;            //接收数据包序号
        private DateTime tx_time = new DateTime();   //发送升级数据时的时间
        private DateTime rx_time = new DateTime();   //接收到下位机响应的时间
        private static int rx_flag=0;               //数据接收标志
        private static long pack_count;              //数据包包数
        private static int pack_size = 256;          //每包包含的固件数据长度
        private static int mcu_ready_ok = 0;         //下位机初始化完成



        //10ms运行一次
        private void timer3_Tick(object sender, EventArgs e)
        {


            if (update_status == 0)                            //未进行升级
            {
                tx_time = System.DateTime.Now;
                return ;
            }

            UpdateStatusBox4.Text = update_status.ToString();   //升级状态

            TimeSpan span = (TimeSpan)(DateTime.Now - tx_time); 

            if ( span.Milliseconds > response_max_time)       //超过response_max_time升级失败
            {
                failed_transmit_cnt++;
                if (failed_transmit_cnt > 3)                  //最多允许失败3次，失败会重新发送升级命令
                {
                    update_status = 0;
                    textBox3.AppendText("\r\n 升级失败");
                }
                else
                {
                    if (update_status == 2)
                    {
                        update_status = 1;
                    }
                    else if (update_status == 3 && rx_flag == 0)
                    {
                        rx_flag = 1;
                        textBox3.AppendText("\r\nGet MCU Response Failed!");
                    }

                    textBox3.AppendText("\r\n 尝试第" + failed_transmit_cnt.ToString() + "次重发。");
                }
                tx_time = System.DateTime.Now;
                return ;
            }


            if (update_status == 1)                          //发送准备请求
            {           
                UpdateReadyInfoSend(sender, e);
                update_status = 2;                          //
            }
            else if (update_status == 2&& mcu_ready_ok == 1)
            {
                textBox3.AppendText("\r\n mcu_ready_ok!\r\n");
                update_status = 3;
                rx_flag = 0;
                mcu_ready_ok = 0;

                UpdateDataSend(sender, e);
   
            }
            else if (update_status == 3&& rx_flag == 1)
            {
                rx_flag = 0;
                UpdateDataSend(sender, e);

            }
            else if (update_status == 4)         //升级完成
            {
                textBox3.AppendText("\r\n 升级完成!");               
                update_status = 0;
                MessageBox.Show("升级完成!");
            }


        }


        OpenFileDialog dialog = new OpenFileDialog();
        private void OpenFile_Click(object sender, EventArgs e)
        {
            /* 弹出文件选择框供用户选择 */
            //OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;                     //该值确定是否可以选择多个文件
            dialog.Title = "请选择要加载的文件(文本格式)";
            dialog.Filter = "Bin文件(*.bin)|*.bin|文本文件(*.txt)|*.txt";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                FilePath.Text = dialog.FileName;
                fs = new FileStream(dialog.FileName, FileMode.Open, FileAccess.Read);      //实例化FileStream
                pack_count = (fs.Length - 1) / pack_size + 1;                             //计算FileStream需要发送多少次  
                SumPackCount.Text = pack_count.ToString();
                byte[] bytes = File.ReadAllBytes(dialog.FileName);

                Crc16(bytes, (int)fs.Length);
                textBox2.Text = "all_crc=" + Crc16(bytes, (int)fs.Length).ToString();
            }
            else
            {
                return;
            }
        }


        private void SendFile_Click(object sender, EventArgs e)
        {

   
            rx_time = System.DateTime.Now;                          //获取当前时间
            update_status = 1;
        }

        private void UpdateReadyInfoSend(object sender, EventArgs e)
        {
            mcu_ready_ok = 0;
            tx_index = 0;

            if (Convert.ToUInt32(MessageHeadBox.Text, 16) > 65535)                    //确认消息头是否合法
            {
                MessageBox.Show("消息头范围为 0x0000~0xFFFF", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            byte[] bArr = new byte[pack_size + 7]; //创建一个Size+7字节的byte[] 其中数据头2字节 序号1字节  数据长度2字节 校验2字节
            UInt16 date = (UInt16)(Convert.ToUInt32(MessageHeadBox.Text, 16));
            UInt16 data_len =4;                                     //升级准备数据帧的数据长度
            bArr[0] = (Byte)((date >> 8) & 0xff);                   //数据头
            bArr[1] = (Byte)(date & 0xff);                          //数据头
            bArr[2] = (Byte)(tx_index);                             //消息序号
            bArr[3] = (Byte)(tx_index);                             //消息序号
            bArr[4] = (Byte)(data_len & 0xFF);                      //数据长度低位
            bArr[5] = (Byte)((data_len >> 8) & 0xFF);               //数据长度高位
            bArr[6] = (Byte)(fs.Length & 0xFF);                      //数据总字节数低位
            bArr[7] = (Byte)((fs.Length >> 8) & 0xFF);               // 
            bArr[8] = (Byte)((fs.Length >> 16) & 0xFF);              //
            bArr[9] = (Byte)((fs.Length >> 24) & 0xFF);              // 数据总字节数高位

            ushort crc_result = Crc16(bArr, data_len + 6);
            bArr[data_len + 6] = (byte)(crc_result & 0xff);             // CRC
            bArr[data_len + 7] = (byte)((crc_result >> 8) & 0xff);      //CRC

            serialPort1.Write(bArr, 0, data_len + 8);
            rx_flag = 0;                                            //数据已经发送
            tx_time = System.DateTime.Now;                          //获取当前时间
            textBox3.AppendText("\r\n[" + tx_time.ToString("HH:mm:ss:fff") + "]" + ":发送升级准备消息 " );
            //extBox3.AppendText("\r\n " +  "CRC="+ crc_result.ToString());

            send_count += bArr.Length;                                   //计数变量累加
            label7.Text = "发送：" + send_count.ToString() + " Bytes";   //刷新界面
    
        }

        private void UpdateDataSend(object sender, EventArgs e)
        {


            if (Convert.ToUInt32(MessageHeadBox.Text, 16) > 65535)                    //确认消息头是否合法
            {
                MessageBox.Show("消息头范围为 0x0000~0xFFFF", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            byte[] bArr = new byte[pack_size + 8]; //创建一个Size+7字节的byte[] 其中数据头2字节 序号1字节  数据长度2字节 校验2字节
            UInt16 date = (UInt16)(Convert.ToUInt32(MessageHeadBox.Text, 16));
            UInt16 read_data_len = 0;
            bArr[0] = (Byte)((date >> 8) & 0xff);//数据头
            bArr[1] = (Byte)(date & 0xff);       //数据头

            if(failed_transmit_cnt==0)  tx_index++;   //发送失败不会重新发送

            if (tx_index != pack_count+1)            //判断是否最后一次发送数据块
            {

                bArr[2] = (Byte)(tx_index );            //消息序号从1开始   
                bArr[3] = (Byte)((tx_index>>8)&0xff);   //消息序号高位

                fs.Position = pack_size * (tx_index-1);                            //更新读取的位置
                read_data_len = (ushort)fs.Read(bArr, 6, pack_size);              //每次读入size 字节到发送缓存
                bArr[4] = (Byte)(read_data_len & 0xFF);                      //数据长度低位
                bArr[5] = (Byte)((read_data_len >> 8) & 0xFF);               //数据长度高位

                ushort crc_result = Crc16(bArr, read_data_len + 6);
                bArr[read_data_len + 6] = (byte)(crc_result & 0xff);             // CRC
                bArr[read_data_len + 7] = (byte)((crc_result >> 8) & 0xff);      //CRC

                serialPort1.Write(bArr, 0, read_data_len + 8);
                rx_flag = 0;
                tx_time = System.DateTime.Now;                          //获取当前时间
                textBox3.AppendText("\r\n[" + tx_time.ToString("HH:mm:ss:fff") + "]" + ":发送第 " + tx_index.ToString() + " 包数据");
                //extBox3.AppendText("\r\n " +  "CRC="+ crc_result.ToString());

                send_count += bArr.Length;                                   //计数变量累加
                label7.Text = "发送：" + send_count.ToString() + " Bytes";   //刷新界面
               
            }
            else
            {
                MessageBox.Show("File Transmit Over!");
            }

        }
        private void button7_Click(object sender, EventArgs e)
        {
            UpdateDataSend(sender,e);
        }

        int i =1;
        private void button4_Click(object sender, EventArgs e)
        {
          
            tx_index = 0;
            i++;
            if (i % 2 == 0)
                timer3.Stop();
            else
                timer3.Start();
        }



    
        // Name: CRC-16/MODBUS    x16+x15+x2+1
        public static ushort Crc16(byte[] buffer, int len )
        {

            ushort crc = 0xFFFF;// Initial value
            for (int i = 0; i<len; i++)
            {
                crc ^= buffer[i];
                for (int j = 0; j< 8; j++)
                {
                    if ((crc & 1) > 0)
                        crc = (ushort) ((crc >> 1) ^ 0xA001);// 0xA001 = reverse 0x8005 
                    else
                        crc = (ushort) (crc >> 1);
                }
            }            

            return crc;
        }
 

   
    }
}
