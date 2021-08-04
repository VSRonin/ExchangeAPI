#ifndef ManagedInterface_h__
#define ManagedInterface_h__
#include <string>
#include <vector>
#using <System.dll>
namespace ManagedInterface {
    bool saveEmail(const std::string& from, const std::vector<std::string>& to, const std::vector<std::string>& cc, const std::vector<std::string>& bcc, const std::string& destination, const std::string& subject, const std::string& body, const std::vector<std::string>& attachments);
    bool sendEmail(const std::string& from, const std::vector<std::string>& to, const std::vector<std::string>& cc, const std::vector<std::string>& bcc, const std::string& host, const std::string& subject, const std::string& body, const std::vector<std::string>& attachments);
    System::String^ getSID();
    System::String^ getUserName();
    bool subscribeEmailReceived();
    bool scanExistingVCONs(int limit);
}
#endif // ManagedInterface_h__