using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CladN
{
    internal class Request
    {
        public int RowNumber;
        public int Code { get; set; }
        public int CodeProduct { get; set; }
        public int CodeClient { get; set; }
        public int RequestNubmer { get; set; }
        public int CountNeeded { get; set; }
        public DateTime DatePost { get; set; }

        public readonly int Id;

        public Request(int id, int code, int codeProduct, int codeClient, int requestNubmer, int countNeeded, DateTime datePost) //datePost to DateTime
        {

            Id = id;
            Code = code;
            CodeProduct = codeProduct;
            CodeClient = codeClient;
            RequestNubmer = requestNubmer;
            CountNeeded = countNeeded;
            DatePost = datePost;
        }

        public void SetCode(int code)
        {
            Code = code;
        }
        public void SetCodeProduct(int codeProduct)
        {
            CodeProduct = codeProduct;
        }
        public void SetCodeClient(int codeClient)
        {
            CodeClient = codeClient;
        }
        public void SetRequestNubmer(int requestNubmer)
        {
            RequestNubmer = requestNubmer;
        }
        public void SetCountNeeded(int countNeeded)
        {
            CountNeeded = countNeeded;
        }
        public void SetDatePost(DateTime datePost)
        {
            DatePost = datePost;
        }
    }
}
