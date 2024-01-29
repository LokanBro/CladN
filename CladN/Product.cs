
namespace CladN
{
    public class Product
    {
        public int Code { get;}    
        public string Name { get;} 
        public double Price { get;}  
        public string Unit { get;}

        public readonly int Id;

        public Product(int id, int code, string name, string unit, double price) 
        {
            Id = id;
            Code = code;
            Name = name;
            Unit = unit;
            Price = price;
        }

        public override bool Equals(object? obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != typeof(Product)) return false;
            return Equals((Product)obj);
        }
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}
