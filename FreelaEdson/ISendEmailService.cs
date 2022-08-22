namespace FreelaEdson
{
    public interface ISendEmailService
    {
        void SendEmail(string email, string subject, string message);
    }
}
