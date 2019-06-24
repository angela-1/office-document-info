using Newtonsoft.Json;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace office_document_info
{


    class DocumentInfo
    {

        #region property
        private string _title = "";
        private string _code = "";
        private string _sendBy = "";
        private string _sendTo = "";
        private string _sendDate = "";
        #endregion

        #region private methods

        private List<string> GetParagraphs(string filePath)
        {
            List<string> paragraphs = new List<string>();
            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                XWPFDocument document = new XWPFDocument(file);
                var pars = document.Paragraphs;
                foreach (var paragraph in document.Paragraphs)
                {
                    paragraphs.Add(paragraph.Text.Trim());
                }
            }
            return paragraphs;
        }

        private string GetCode(string paragraph)
        {
            string value = "";
            Regex reg = new Regex(@"\S+〔\d{4}〕\d+号");
            Match match = reg.Match(paragraph);
            if (match.Success)
            {
                value = match.Value;
            }
            return value;
        }

        private string GetSendTo(string paragraph)
        {
            string value = "";
            Regex reg = new Regex(@"^\S+[：:]$");
            Match match = reg.Match(paragraph);
            if (match.Success)
            {
                value = match.Value;
            }
            return value;
        }

        private string GetSendDate(string paragraph)
        {
            string value = "";
            Regex reg = new Regex(@"^\d{4}年\d{1,2}月\d{1,2}日$");
            Match match = reg.Match(paragraph);
            if (match.Success)
            {
                value = match.Value;
            }
            return value;
        }

        private bool IsWhiteLine(string paragraph)
        {
            Regex reg = new Regex(@"^\s*$");
            Match match = reg.Match(paragraph);
            return match.Success;
        }


        private bool Parse(List<String> contents)
        {
            // 标记各值是否取得
            // 0b0001 文号
            // 0b0010 标题
            // 0b0100 主送
            // 0b1000 发文日期
            int flag = 0b0000;

            const int HAS_CODE = 1;
            const int HAS_TITLE = 2;
            const int HAS_SEND_TO = 4;
            const int HAS_SEND_DATE = 8;

            bool hasTitle = false;

            foreach (var line in contents)
            {
                if ((flag & HAS_CODE) == 0 && (flag & HAS_TITLE) == 0)
                {
                    string code = GetCode(line);
                    if (code.Length > 0)
                    {
                        _code = code;
                        flag |= 1;
                        continue;
                    }
                }

                if ((flag & HAS_SEND_TO) == 0)
                {
                    string sendTo = GetSendTo(line);
                    if (sendTo.Length > 0)
                    {
                        int ind = contents.IndexOf(line);
                        List<string> titleArray = new List<string>();
                        for (int i = 1; i <= ind; i++)
                        {
                            string t = contents[ind - i];
                            titleArray.Add(t);
                            if (t.Length > 0)
                            {
                                hasTitle = true;
                            }
                            if (IsWhiteLine(t) && hasTitle)
                            {
                                titleArray.Reverse();
                                _title = string.Join("", titleArray);
                                flag |= 2;
                                break;
                            }
                        }
                        _sendTo = sendTo;
                        flag |= 4;
                        continue;
                    }
                }

                if ((flag & HAS_SEND_DATE) == 0)
                {
                    string sendDate = GetSendDate(line);
                    if (sendDate.Length > 0)
                    {
                        int ind = contents.IndexOf(line);
                        _sendBy = contents[ind - 1];
                        _sendDate = sendDate;
                        flag |= 8;
                        continue;
                    }
                }

                if (flag == 0b1111)
                {
                    break;
                }
            }
            return true;
        }


        #endregion


        #region public methods
        public string GetInfo(string filePath, OutputFormat format)
        {
            List<string> paragraphs = GetParagraphs(filePath);
            string result = "";
            if (paragraphs.Count > 0)
            {
                Parse(paragraphs);
                if (format == OutputFormat.JSON)
                {
                    result = JsonConvert.SerializeObject(
                        new { title = _title, code = _code, sendBy = _sendBy, sendTo = _sendTo, sendDate = _sendDate });
                }
                else
                {
                    result = _sendDate + "\t" + _sendBy + "\t" + _code + "\t" + _title;
                }
            }
            return result;
        }

        #endregion
    }
}
