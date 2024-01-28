
namespace CladN
{
    public class Request
    {
        public int RowNumber;
        public int Code { get;}
        public int CodeProduct { get;}
        public int CodeClient { get;}
        public int RequestNubmer { get;}
        public int CountNeeded { get;}
        public DateTime DatePost { get;}

        public readonly int Id;

        public Request(int id, int code, int codeProduct, int codeClient, int requestNubmer, int countNeeded, DateTime datePost)
        {

            Id = id;
            Code = code;
            CodeProduct = codeProduct;
            CodeClient = codeClient;
            RequestNubmer = requestNubmer;
            CountNeeded = countNeeded;
            DatePost = datePost;
        }
    }
}
