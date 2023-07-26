using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Data;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Diagnostics;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using System.ServiceProcess;
using System.Configuration;
using System.Runtime.InteropServices;
using PDF_Common;
using NLog;
namespace pdf_text
{
    class Program
    {
        static string txt_path = Common.Read_Path("txt_path"); 
        static string log_path = Common.Read_Path("log_path"); 
        static string pdf_path = Common.Read_Path("pdf_path"); 
        static string sample_path = Common.Read_Path("sample_path"); 
        static void Main(string[] args)
        {
            Thread[] myThread = new Thread[args.Length];
            int k;
            for (k = 0; k < args.Length; k++)
            {
                myThread[k].Start(txt_path + args[k]); 
            }
        }

        static void process(object fFileName)
        {
            string action = "";
            string printer = "";
            string page_size = "";
            string page_write = "";
            string word_write = "";
            string sample_name = "";
            string ftp_ip = "";
            string ftp_id = "";
            string ftp_pwd = "";
            string ftp_path = "";
            string PDF_FILE = "";
            string txt_Name = "";
            string encry = "";
            ThreadPriority priority = new ThreadPriority();
            string flag_Name = fFileName as string; 
            string subdirectory = "";
            string prttray = ""; 

            float top_dist = 20; 
            float column_dist = 11; 

            try
            {
                if (flag_Name[flag_Name.LastIndexOf("\\") + 1] == '7') 
                {
                    subdirectory = @"U40\";
                }
                else if (flag_Name[flag_Name.LastIndexOf("\\") + 1] == '8') 
                {
                    subdirectory = @"VPN\";
                }
                else
                {
                    subdirectory = "";
                }

                if (System.IO.File.Exists(flag_Name))
                {
                    Common.Read_flag(flag_Name, ref action, ref printer, ref page_size, ref page_write, ref word_write, ref sample_name,
                                 ref ftp_ip, ref ftp_id, ref ftp_pwd, ref ftp_path, ref txt_Name, ref PDF_FILE, ref priority, ref encry, ref top_dist, ref column_dist, ref prttray);
                }
                else
                {
                    Logger logger = LogManager.GetCurrentClassLogger();
                    logger.Error(string.Format("exception 讀取{0}發生錯誤，找不到{1}", fFileName));
                    Console.WriteLine(string.Format("exception 讀取{0}發生錯誤，找不到{1}", fFileName));
                }
            }
            catch (IOException ioex)
            {
                Logger logger = LogManager.GetCurrentClassLogger();
                logger.Error(string.Format("exception 讀取{0}發生錯誤{1}", fFileName, ioex.ToString()));
                Console.WriteLine(string.Format("exception 讀取{0}發生錯誤{1}", fFileName, ioex.ToString()));
                File.Move(flag_Name, flag_Name.Replace(".log", ".flag"));
                return;
            }
            catch (Exception ex)
            {
                Logger logger = LogManager.GetCurrentClassLogger();
                logger.Error("pdf_text Common.Read_flag()異常:" + ex.ToString());
                Console.WriteLine("pdf_text Common.Read_flag()異常:" + ex.ToString());
                return;
            }

            if (flag_Name.ToLower().Contains("~part"))
            {
                action = "file";
            }

            Thread.CurrentThread.Priority = priority; 
            try 
            {
                if (!System.IO.Directory.Exists(pdf_path + subdirectory))
                {
                    System.IO.Directory.CreateDirectory(pdf_path + subdirectory);
                }

                if (!System.IO.Directory.Exists(log_path + subdirectory))
                {
                    System.IO.Directory.CreateDirectory(log_path + subdirectory);
                }

                if ((txt_Name[0] == '7' || txt_Name[0] == '8') && txt_Name.ToUpper().EndsWith(".MSG"))
                {
                    File.Copy(txt_path + txt_Name, pdf_path + subdirectory + txt_Name, true);
                    File.Delete(txt_path + txt_Name);
                }
                else
                {
                    StreamReader file = new System.IO.StreamReader(txt_path + txt_Name, System.Text.Encoding.UTF8);
                    C_PDF(file, action, printer, page_size, page_write, word_write, sample_name, ftp_ip, ftp_id, ftp_pwd, ftp_path, encry, PDF_FILE, subdirectory, top_dist, column_dist, prttray); //產生PDF報表並印到印表機, FTP
                }
                if (flag_Name.ToLower().Contains("~part"))
                {   
                    File.Move(pdf_path + subdirectory + txt_Name + ".pdf", (pdf_path + subdirectory + txt_Name + ".pdf").Replace("~", "-"));
                    File.Delete(flag_Name);
                    File.Delete(flag_Name.Replace(".log", ".error"));
                    File.Delete(flag_Name.Substring(0, flag_Name.LastIndexOf(".")));
                }
                else
                {   
                    File.Delete(log_path + subdirectory + flag_Name.Substring(flag_Name.LastIndexOf("\\") + 1).ToLower().Replace(".log", ".error"));
                    File.Delete(log_path + subdirectory + flag_Name.Substring(flag_Name.LastIndexOf("\\") + 1));
                    File.Move(flag_Name, log_path + subdirectory + flag_Name.Substring(flag_Name.LastIndexOf("\\") + 1));
                }
            }
            catch (WebException ex) 
            {   
                StreamWriter sw = new StreamWriter(File.Open(flag_Name, FileMode.Append), System.Text.Encoding.Default);
                sw.Write("\n" + "pdf_text process()異常" + ex.ToString());
                sw.Close();
                sw.Dispose();

                File.Delete(log_path + subdirectory + flag_Name.Substring(flag_Name.LastIndexOf("\\") + 1));
                File.Delete(log_path + subdirectory + flag_Name.Substring(flag_Name.LastIndexOf("\\") + 1).ToLower().Replace(".log", ".error"));
                File.Move(flag_Name, log_path + subdirectory + flag_Name.Substring(flag_Name.LastIndexOf("\\") + 1).ToLower().Replace(".log", ".error"));
                File.Delete(pdf_path + subdirectory + @"FTPerror\" + PDF_FILE);
                File.Move(pdf_path + subdirectory + PDF_FILE, pdf_path + subdirectory + @"FTPerror\" + PDF_FILE);
            }
            catch (Exception ex)  
            {
                StreamWriter sw = new StreamWriter(File.Open(flag_Name, FileMode.Append), System.Text.Encoding.Default);
                sw.Write("\n" + "pdf_text process()異常:" + ex.ToString());
                sw.Close();
                sw.Dispose();

                File.Delete(log_path + subdirectory + flag_Name.Substring(flag_Name.LastIndexOf("\\") + 1));
                File.Delete(log_path + subdirectory + flag_Name.Substring(flag_Name.LastIndexOf("\\") + 1).ToLower().Replace(".log", ".error"));
                File.Move(flag_Name, log_path + subdirectory + flag_Name.Substring(flag_Name.LastIndexOf("\\") + 1).ToLower().Replace(".log", ".error"));
            }
        }

        
        static void C_PDF(StreamReader file, string action, string printer, string page_size, string page_write, string word_write, string sample_name, string ftp_ip,
                          string ftp_id, string ftp_pwd, string ftp_path, string encry, string PDF_FILE, string subdir, float top_dist, float column_dist, string prttray)
        {
            Logger logger = LogManager.GetCurrentClassLogger();
            logger.Debug("pdf_text開始轉置:" + PDF_FILE);
            Console.WriteLine("pdf_text開始轉置:" + PDF_FILE);

            DateTime dteS = DateTime.Now;
            Rectangle rpt_size = Common.get_pagesize(page_size, page_write);
            Common common = new Common();
            Font font = common.get_font(""); 
            Chunk chunk; 
            Phrase phrase; 
            Phrase tmp_phrase = new Phrase();
            Paragraph paragraph = new Paragraph(); 
            paragraph.SetLeading(column_dist, 0);
            PdfReader reader;
           Barcode barcode = new Barcode39();

            BarcodeQRCode qrcode; 
            string qrcode_text = "";

            string _str; 
            string tmp_pcl = ""; 
            int NewPage_flag = 0; 
            float WordSpace = 0; 
            int WordWidth = 1; 
            int line_cnt = 0;
            int ctrl_LN = 0;  
            int next_line_flag = 0;
            int dtl_morepage_flag = 0; 
            Document doc = new Document(rpt_size, 15, 0, top_dist, 0); 
            File.Delete(pdf_path + subdir + PDF_FILE);
            PdfWriter docwriter = PdfWriter.GetInstance(doc, new FileStream(pdf_path + subdir + PDF_FILE, FileMode.CreateNew));

            try
            {
                if (encry != "")
                {  
                    docwriter.SetEncryption(PdfWriter.DO_NOT_ENCRYPT_METADATA, encry, encry, PdfWriter.ALLOW_PRINTING | PdfWriter.ALLOW_COPY);
                }
                doc.Open();


                if (sample_name != "") 
                {
                    reader = new PdfReader(sample_path + sample_name);
                    PdfImportedPage pi = docwriter.GetImportedPage(reader, 1);
                    docwriter.DirectContentUnder.AddTemplate(pi, 0, 0);
                }

                while ((_str = file.ReadLine()) != null)
                {
                    int esc_count = 0;
                    int FF_location = -1;
                    int bar_on = 0;
                    int bar_long = 0;
                    int qr_on = 0;      
                    int qr_long = 0;
                    float ctrl_LH = 0;  
                    float ctrl_LH_mk = 0, MAX_FontSize_Phrase = common.font_size;  

                    _str = _str.TrimEnd(new char[] { ' ', '\r', '\n' });
                    if (line_cnt == 0)
                        _str = _str.Replace("\f", "");

                    line_cnt++;
                    if (_str.Trim().Contains("此聯空白作廢") == true && (sample_name.ToLower() == "sample0.pdf" || sample_name.ToLower() == "firstpage_a4.pdf"))
                        dtl_morepage_flag = 1;

                    if (_str.Trim().Contains("報") == true || _str.Trim().Contains("表") == true)
                        next_line_flag = 1;
                    else if (line_cnt == 1 && _str.Trim() == "")
                        next_line_flag = 0;

                    if ((line_cnt >= 1 && line_cnt <= 2) && (sample_name.ToLower() == "sample0.pdf" || sample_name.ToLower() == "firstpage_a4.pdf") && next_line_flag == 0)
                    {
                        continue;
                    }

                    if (sample_name.ToLower() == "sample0.pdf" || sample_name.ToLower() == "sample1.pdf")
                    {
                        if (line_cnt == 22 && next_line_flag == 0)
                        {
                            phrase = new Phrase();
                            phrase.Add(new Chunk("\n\n\n\n"));
                            phrase.SetLeading(column_dist, 0);
                            doc.Add(phrase);
                            phrase.Clear();
                        }
                        if (line_cnt <= 46 && dtl_morepage_flag == 1 && _str.Trim().Contains("WKAIFS8") == true)
                        {
                            phrase = new Phrase();
                            for (int i = 46; i > line_cnt; i--) 
                            {
                                phrase.Add(new Chunk("\n"));
                            }
                            phrase.SetLeading(column_dist, 0);
                            doc.Add(phrase);
                            phrase.Clear();
                        }
                    }
                    byte[] _str_asc = Encoding.ASCII.GetBytes(_str);
                    phrase = new Phrase();
                    phrase.SetLeading(column_dist, 0); //line-height(fix, multi)
                    tmp_phrase.SetLeading(column_dist, 0);
                    if (NewPage_flag == 1) 
                    {

                        if (file.Peek() < 0) 
                        {
                            for (int j = 0; j < _str_asc.Length; j++)
                            {
                                if (_str_asc[j] == 27)
                                {
                                    esc_count++;
                                }
                                else if (_str_asc[j] == 12)
                                {
                                    FF_location = j;
                                    esc_count = 0;
                                }
                            }
                            if (esc_count * 5 == _str.Length || _str.Substring(FF_location + 1, _str.Length - FF_location - 1).Length == esc_count * 5)
                            {
                                continue;
                            }
                            else
                            {
                                doc.SetMargins(15, 0, top_dist, 0);
                                doc.NewPage();
                            }
                        }
                        else
                        {
                           doc.SetMargins(15, 0, top_dist - column_dist, 0);

                            doc.NewPage();


                            if (tmp_phrase.Count > 0)
                            {
                                doc.Add(tmp_phrase);
                                tmp_phrase.Clear();
                            }
                        }
                        if (sample_name != "")
                        {
                            reader = new PdfReader(sample_path + sample_name);
                            PdfImportedPage pi = docwriter.GetImportedPage(reader, 1);
                            docwriter.DirectContentUnder.AddTemplate(pi, 0, 0);
                        }
                        NewPage_flag = 0;
                    }
                    for (int i = 0; i < _str_asc.Length; i++)
                    {
                        if (_str_asc[i] == 12) 
                        {
                            NewPage_flag = 1;
                            line_cnt = 0;
                            continue;
                        }

                        if (_str_asc[i] == 27) 
                        {
                            if (i + 1 + 4 <= _str.Length)
                            {
                                tmp_pcl = _str.Substring(i + 1, 4);
                            }

                            if (_str[i + 1].ToString().ToUpper() == "B") 
                            {
                                i += 4;
                                if (Common.get_barcode(tmp_pcl) == 99) 
                                {
                                    bar_on = 1;
                                    barcode.Code = "";
                                }
                                else if (Common.get_barcode(tmp_pcl) == 1) 
                                {
                                    PdfContentByte cb = docwriter.DirectContent;
                                    barcode.AltText = barcode.Code; 
                                    if (barcode.CodeType == 0)
                                    {
                                        barcode.AltText = ""; 
                                        barcode.Size = font.Size * 72 / 96;
                                        barcode.Baseline = barcode.Size;
                                        if (PDF_FILE.Substring(1, 7) == "f05811r")
                                        {
                                            barcode.Size = font.Size * 62 / 96;
                                            barcode.Baseline = barcode.Size;
                                            barcode.BarHeight = -2.5f * font.Size;
                                            barcode.X = 0.772f; 
                                        }
                                        else
                                        {
                                            barcode.BarHeight = -2.5f * font.Size;
                                            barcode.X = 0.672f; 
                                        }
                                    }
                                    chunk = new Chunk(barcode.CreateImageWithBarcode(cb, null, null), 0, barcode.Baseline, true);
                                    if (NewPage_flag == 0)
                                    {
                                        phrase.Add(chunk);
                                    }
                                    else
                                    {
                                        tmp_phrase.Add(chunk);
                                    }
                                    bar_on = 0;
                                    barcode = new Barcode39();
                                }
                                else if (Common.get_barcode(tmp_pcl) == 128) 
                                    barcode = new Barcode128();
                                }
                            }
                            else if (_str[i + 1] == 'Q') 
                            {
                                i += 4;
                                if (Common.get_barcode(tmp_pcl) == 98) 
                                {
                                    qr_on = 1;
                                    qrcode_text = "";
                                }
                                else if (Common.get_barcode(tmp_pcl) == 2) 
                                {
                                    qrcode = new BarcodeQRCode(qrcode_text.Trim(), 0, 0, null);

                                    PdfContentByte cb = docwriter.DirectContent;
                                    Image vImg = qrcode.GetImage();
                                   vImg.ScaleAbsoluteHeight(69);
                                    vImg.ScaleAbsoluteWidth(69);
                                    float phrase_width = 0;
                                    foreach (Chunk vchunk in phrase.Chunks)
                                    {
                                        phrase_width += vchunk.GetWidthPoint();
                                    }
                                    vImg.SetAbsolutePosition(phrase_width - 22, docwriter.GetVerticalPosition(false) - 30);
                                    cb.AddImage(vImg);
                                    qr_on = 0;

                                }
                                else if (Common.get_barcode(tmp_pcl) == 128) 
                                {
                                    barcode = new Barcode128();
                                }
                            }
                            else if (_str[i + 1] == 'r' || tmp_pcl.ToUpper() == "000E") 
                            {
                                i += 4;
                                WordSpace = common.get_wd(tmp_pcl, page_size, page_write); 
                            }
                            else if (_str.Substring(i + 1, 2).ToUpper() == "LN") 
                            {
                                i += 4;
                                ctrl_LN = Convert.ToInt16(tmp_pcl.Substring(2));
                            }
                            else if (_str.Substring(i + 1, 2).ToUpper() == "LH") 
                            {
                                i += 4;
                                ctrl_LH_mk = 1;
                                ctrl_LH = Convert.ToSingle(tmp_pcl.Substring(2));
                            }
                            else
                            {
                                i += 4;
                                font = common.get_font(tmp_pcl);
                                WordWidth = common.ww; 
                                MAX_FontSize_Phrase = common.font_size > MAX_FontSize_Phrase ? common.font_size : MAX_FontSize_Phrase;
                            }
                        }
                        else if (bar_on == 1)
                        {
                            bar_long += 1;
                            barcode.Code += _str.Substring(i, 1);
                        }
                        else if (qr_on == 1)   
                        {
                            qr_long += 1;
                            qrcode_text += _str.Substring(i, 1);
                        }
                        else 
                        {
                            chunk = new Chunk(Substr(_str, i, 1), font); 
                            if (Encoding.Default.GetByteCount(chunk.Content) == 1) 
                            {
                                chunk.SetCharacterSpacing(WordSpace / 2);
                            }
                            else
                            {
                                chunk.SetCharacterSpacing(WordSpace);
                            }
                            chunk.SetHorizontalScaling(WordWidth);
                            //phrase.Add(chunk);
                            if (NewPage_flag == 0)
                            {
                                phrase.Add(chunk);
                            }
                            else
                            {
                                tmp_phrase.Add(chunk);
                            }
                        }
                    }
                    if (NewPage_flag == 0)
                    {
                        phrase.Add(new Chunk("\n", font));
                        tmp_phrase.Clear();
                    }
                    else
                    {
                        tmp_phrase.Add(new Chunk("\n", font));
                    }
                    //}

                    if (ctrl_LH_mk == 1)
                    {
                        ctrl_LH_mk = 0;
                        phrase.SetLeading(ctrl_LH, 0);
                    }
                    else
                    {
                        phrase.SetLeading(common.get_column_dist(page_size, page_write, MAX_FontSize_Phrase, top_dist, ctrl_LN), 0); 
                    }
                    doc.Add(phrase); 
                    phrase.Clear();
                }

                tmp_phrase.Clear();
                doc.Close();
                doc.Dispose();
                file.Close();
                file.Dispose();
                docwriter.Dispose();

                TimeSpan ts = DateTime.Now.Subtract(dteS);
                logger.Debug("pdf_text轉置完成:" + PDF_FILE + " 耗時:" + ts + "(秒)");
                Console.WriteLine(@"pdf_text轉置完成:" + PDF_FILE + " 耗時:" + ts + "(秒)");

                switch (action.ToLower())
                {
                    case "print": 
                        if (prttray != "" && prttray != null) 
                            Common.SendToPrinterViaGSBATCHPRINT(printer, PDF_FILE, page_size, page_write, prttray);
                        else
                            Common.SendToPrinterViaGSBATCHPRINT(printer, PDF_FILE, page_size, page_write); 
                        Console.WriteLine(@"pdf_text轉置成功，已順利" + action.ToLower() + " " + PDF_FILE + " printer=" + printer);
                        logger.Debug("pdf_text轉置成功，已順利" + action.ToLower() + " " + PDF_FILE + " printer=" + printer);
                        break;

                    case "file":  
                        Console.WriteLine(@"pdf_text轉置成功，已順利" + action.ToLower() + " " + PDF_FILE);
                        logger.Debug("pdf_text轉置成功，已順利" + action.ToLower() + " " + PDF_FILE);
                        break;
                    case "ftp":   
                        string _result = ""; 
                        if (PDF_FILE[0] == '7') 
                        {
                            Common.ftpfile_rename(ftp_path, pdf_path + subdir + PDF_FILE, PDF_FILE.Substring(8), ftp_ip, ftp_id, ftp_pwd, ref _result);
                        }
                        else if (PDF_FILE[0] == '8') 
                        {
                            Common.ftpfile_rename(ftp_path, pdf_path + subdir + PDF_FILE, PDF_FILE.Substring(8), ftp_ip, ftp_id, ftp_pwd, ref _result);
                        }
                        else
                        {
                            Common.ftpfile(ftp_path, pdf_path + subdir + PDF_FILE, ftp_ip, ftp_id, ftp_pwd, ref _result);
                        }
                        Console.WriteLine(@"成功" + action.ToLower() + " " + ftp_ip + ftp_path + PDF_FILE + " result=" + _result);
                        logger.Debug("成功" + action.ToLower() + " " + ftp_ip + ftp_path + PDF_FILE + " result=" + _result);
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(@"異常:" + PDF_FILE + "訊息" + ex.ToString());
                logger.Debug("異常:" + PDF_FILE + "訊息" + ex.ToString());
            }
        }


        static void Addcolumn(Rectangle rpt_size, ref PdfWriter docwriter, ref int page, ref Paragraph paragraph)
        {
            try
            {
                page++;
                PdfContentByte write = docwriter.DirectContent;
                ColumnText column = new ColumnText(write);
                column.SetSimpleColumn(rpt_size);
                column.Indent = 40;
                column.SetLeading(2f, 1);
                column.AddText(paragraph);
                column.Go();
                paragraph.Clear();
            }
            catch (Exception ex)
            {
                throw;
            }

        }

        public static int getchar(string s)
        {
            try
            {
                return Encoding.UTF32.GetCharCount(Encoding.UTF32.GetBytes(s));
            }
            catch (Exception ex)
            {
                throw;
            }

        }
        public static int Len(string s) 
        {
            try
            {
                return Encoding.UTF32.GetByteCount(s) / 4;
            }
            catch (Exception ex)
            {
                throw;
            }

        }
        public static string Substr(string s, int startIndex, int length)
        {
            try
            {
                byte[] byte32Array = Encoding.UTF32.GetBytes(s);
                startIndex *= 4;
                length *= 4;

                if (startIndex >= byte32Array.Length)
                {
                    return "";
                }

                length = (startIndex + length) > byte32Array.Length ? byte32Array.Length - startIndex : length;

                return Encoding.UTF32.GetString(byte32Array, startIndex, length);
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}