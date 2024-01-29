
namespace CladN
{
    public class Request
    {
        public int RowNumber;
        public int Code { get;}
        public Product Product { get;}
        public Client Client { get;}
        public int RequestNubmer { get;}
        public int CountNeeded { get;}
        public DateTime DatePost { get;}

        public readonly int Id;

        public Request(int id, int code, Product product, Client client, int requestNubmer, int countNeeded, DateTime datePost)
        {

            Id = id;
            Code = code;
            Product = product;
            Client = client;
            RequestNubmer = requestNubmer;
            CountNeeded = countNeeded;
            DatePost = datePost;
        }
    }
}
