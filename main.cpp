#include "ManagedInterface.h"
#include <iostream>
#include <string>
#include <sstream>
#ifdef _DEBUG
#define DebugOut(x) std::cerr << std::endl << x;
#else
#define DebugOut(x)
#endif // _DEBUG

enum ReturnValues
{
    NoErrors = 0,
    // Errors
    NoCommand = -1,
    InvalidInput = 1,
    InvalidParameters = 2,
    FailedCommand = 3,
    // Warnings
    ParameterIgnored = 1024
};

int main()
{
    using std::cin;
    typedef std::stringstream streamer;
    int numArgs = 0;
    std::string test;
    std::getline(cin, test, '\n');
    if (test == "SendEmail") {
        DebugOut("Sending Email");
        std::vector<std::string> toList, ccList, bccList;
        std::string from, subject, body, host, saveDestination;
        std::vector<std::string> attachFiles;
        for (;;) {
            std::getline(cin, test, '\n');
            if (test == "To") {
                std::getline(cin, test, '\n');
                streamer(test) >> numArgs;
                DebugOut(numArgs << " To adresses");
                while (numArgs--) {
                    std::getline(cin, test, '\n');
                    DebugOut("To address: " << test);
                    toList.push_back(test);
                }
            }
            else if (test == "Cc") {
                std::getline(cin, test, '\n');
                streamer(test) >> numArgs;
                DebugOut(numArgs << " Cc adresses");
                while (numArgs--) {
                    std::getline(cin, test, '\n');
                    DebugOut("Cc address: " << test);
                    ccList.push_back(test);
                }
            }
            else if (test == "Bcc") {
                std::getline(cin, test, '\n');
                streamer(test) >> numArgs;
                DebugOut(numArgs << " Bcc adresses");
                while (numArgs--) {
                    std::getline(cin, test, '\n');
                    DebugOut("Bcc address: " << test);
                    bccList.push_back(test);
                }
            }
            else if (test == "From") {
                std::getline(cin, from, '\n');
                DebugOut("From address: " << from);
            }
            else if (test == "Subject") {
                std::getline(cin, subject, '\n');
                DebugOut("Subject: " << subject);
            }
            else if (test == "Body") {
                std::getline(cin, body, '\n');
                DebugOut("Body: " << body);
            }
            else if (test == "Attachments") {
                std::getline(cin, test, '\n');
                streamer(test) >> numArgs;
                DebugOut(numArgs << " Attachments");
                while (numArgs--) {
                    std::getline(cin, test, '\n');
                    DebugOut("Attachment: " << test);
                    attachFiles.push_back(test);
                }
            }
            else if (test == "Host") {
                std::getline(cin, host, '\n');
                DebugOut("Host: " << host);
            }
            else if (test == "SaveDestination") {
                std::getline(cin, saveDestination, '\n');
                DebugOut("Save Destination: " << saveDestination);
            }
            else if (test == "Save") {
                DebugOut("Saving");
                if (!ManagedInterface::saveEmail(from, toList, ccList, bccList, saveDestination, subject, body, attachFiles))
                    return FailedCommand;
            }
            else if (test == "Send") {
                DebugOut("Sending");
                if (!ManagedInterface::sendEmail(from, toList, ccList, bccList, host, subject, body, attachFiles))
                    return FailedCommand;
            }
            else if (test.size() == 0) {
                break;
            }
            else {
                DebugOut("Invalid Input: " << test);
                return InvalidInput;
            }
        }
        return NoErrors;
    }
    else if (test == "GetSID"){
        System::String^ currSID = ManagedInterface::getSID();
        if (currSID->Length == 0)
            return FailedCommand;
        System::Console::Write(currSID);
        return NoErrors;
    }
    else if (test == "GetUserName") {
        System::String^ currSID = ManagedInterface::getUserName();
        if (currSID->Length == 0)
            return FailedCommand;
        System::Console::Write(currSID);
        return NoErrors;
    }
    else if (test == "SubVCON") {
        if(!ManagedInterface::subscribeEmailReceived())
            return FailedCommand;
        return NoErrors;
    }
    else if (test == "ScanVCON"){
        std::getline(cin, test, '\n');
        int scanLimit = 0;
        streamer(test) >> scanLimit;
        if (!ManagedInterface::scanExistingVCONs(scanLimit))
            return FailedCommand;
        return NoErrors;
    }
    return NoCommand;
}
