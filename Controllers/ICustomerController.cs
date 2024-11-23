using Models;
namespace Controllers
{
    public interface ICustomerController
    {
        bool IsNameValid(string name);
        bool IsEmailValid(string name);
        bool IsPhoneValid(string name);

        int CompareCustomerById(Customer customer1, Customer customer2);
        int CompareCustomerByName(Customer customer1, Customer customer2);
        int CompareCustomerByPointDESC(Customer customer1, Customer customer2);
        int CompareCustomerByCreatedDate(Customer customer1, Customer customer2);
        int CompareCustomerByBirthDate(Customer customer1, Customer customer2);
    }
}
