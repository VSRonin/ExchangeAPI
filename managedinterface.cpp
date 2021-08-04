#include "ManagedInterface.h"
#include <vcclr.h>
#include <iostream>
#include <string>
#using <System.DirectoryServices.AccountManagement.dll>
#using <Microsoft.Exchange.WebServices.dll>
#undef GetTempPath

bool SslRedirectionCallback(System::String^ serviceUrl)
{
    System::Uri^ redirectionUri = gcnew System::Uri(serviceUrl);
    return redirectionUri->Scheme == "https";
}
namespace ManagedInterface {

    bool saveEmail(const std::string& from, const std::vector<std::string>& to, const std::vector<std::string>& cc, const std::vector<std::string>& bcc, const std::string& destination, const std::string& subject, const std::string& body, const std::vector<std::string>& attachments)
    {
        using namespace System;
        using namespace System::Text;
        using namespace System::Net::Mail;
        String^ tempFolder = IO::Path::GetTempPath() + '\\' + Guid::NewGuid().ToString();
#ifdef _DEBUG
        System::Console::WriteLine("Temp Folder: " + tempFolder);
#endif
        IO::Directory::CreateDirectory(tempFolder);
        SmtpClient^ smtpClient = gcnew SmtpClient();
        MailMessage^ message = gcnew MailMessage();
        message->From = gcnew MailAddress(gcnew String(from.c_str()));
        for (auto i = std::begin(to); i != std::end(to); ++i)
            message->To->Add(gcnew String(i->c_str()));
        for (auto i = std::begin(cc); i != std::end(cc); ++i)
            message->CC->Add(gcnew String(i->c_str()));
        for (auto i = std::begin(bcc); i != std::end(bcc); ++i)
            message->Bcc->Add(gcnew String(i->c_str()));
        for (auto i = std::begin(attachments); i != std::end(attachments); ++i)
            message->Attachments->Add(gcnew Attachment(gcnew String(i->c_str())));
        message->Subject = gcnew String(subject.c_str());
        message->IsBodyHtml = true;
        message->Body = gcnew String(body.c_str());
        smtpClient->DeliveryMethod = SmtpDeliveryMethod::SpecifiedPickupDirectory;
        smtpClient->PickupDirectoryLocation = tempFolder;
        try {
            smtpClient->Send(message);
            auto fileEntries = IO::Directory::GetFiles(tempFolder, "*.eml", IO::SearchOption::TopDirectoryOnly);
            Collections::IEnumerator^ files = fileEntries->GetEnumerator();
            while (files->MoveNext()) {
                String^ fileName = safe_cast<String^>(files->Current);
                String^ destinationFile = gcnew String(destination.c_str());
                if (IO::File::Exists(destinationFile))
                    IO::File::Delete(destinationFile);
                IO::File::Move(fileName, destinationFile);
                IO::Directory::Delete(tempFolder,true);
            }
        }
#ifdef _DEBUG
        catch (ArgumentNullException^ ex) {
            System::Console::WriteLine("ArgumentNullException: " + ex->Message);
            return false;
        }
        catch (InvalidOperationException^ ex) {
            System::Console::WriteLine("InvalidOperationException: " + ex->Message);
            return false;
        }
        catch (SmtpException^ ex) {
            System::Console::WriteLine("SmtpException: " + ex->Message);
            return false;
        }
        catch (ArgumentException^ ex) {
            System::Console::WriteLine("ArgumentException: " + ex->Message);
            return false;
        }
        catch (UnauthorizedAccessException^ ex) {
            System::Console::WriteLine("UnauthorizedAccessException: " + ex->Message);
            return false;
        }
        catch (IO::PathTooLongException^ ex) {
            System::Console::WriteLine("PathTooLongException: " + ex->Message);
            return false;
        }
        catch (IO::DirectoryNotFoundException^ ex) {
            System::Console::WriteLine("DirectoryNotFoundException: " + ex->Message);
            return false;
        }
        catch (IO::IOException^ ex) {
            System::Console::WriteLine("IOException: " + ex->Message);
            return false;
        }
        catch (NotSupportedException^ ex) {
            System::Console::WriteLine("NotSupportedException: " + ex->Message);
            return false;
        }
#else
        catch (ArgumentNullException^) {
            return false;
        }
        catch (InvalidOperationException^) {
            return false;
        }
        catch (SmtpException^) {
            return false;
        }
        catch (ArgumentException^) {
            return false;
        }
        catch (UnauthorizedAccessException^) {
            return false;
        }
        catch (IO::PathTooLongException^) {
            return false;
        }
        catch (IO::DirectoryNotFoundException^) {
            return false;
        }
        catch (IO::IOException^) {
            return false;
        }
        catch (NotSupportedException^) {
            return false;
        }
#endif // _DEBUG
        return true;
    }
    bool sendEmail(const std::string& from, const std::vector<std::string>& to, const std::vector<std::string>& cc, const std::vector<std::string>& bcc, const std::string& host, const std::string& subject, const std::string& body, const std::vector<std::string>& attachments)
    {
        using namespace System;
        using namespace System::Text;
        using namespace System::Net::Mail;
        if (to.size() == 0 && cc.size() == 0 && bcc.size() == 0)
            return false;
        SmtpClient^ smtpClient = gcnew SmtpClient();
        MailMessage^ message = gcnew MailMessage();
        message->From = gcnew MailAddress(gcnew String(from.c_str()));
        for (auto i = std::begin(to); i != std::end(to); ++i)
            message->To->Add(gcnew String(i->c_str()));
        for (auto i = std::begin(cc); i != std::end(cc); ++i)
            message->CC->Add(gcnew String(i->c_str()));
        for (auto i = std::begin(bcc); i != std::end(bcc); ++i)
            message->Bcc->Add(gcnew String(i->c_str()));
        for (auto i = std::begin(attachments); i != std::end(attachments); ++i) {
            String^ tempString = gcnew String(i->c_str());
            Attachment^ tempAttach = gcnew Attachment(tempString);
            tempAttach->Name = System::IO::Path::GetFileName(tempString);
            message->Attachments->Add(tempAttach);
        }
        message->Subject = gcnew String(subject.c_str());
        message->IsBodyHtml = true;
        message->Body = gcnew String(body.c_str());
        smtpClient->Host = gcnew String(host.c_str());
        smtpClient->UseDefaultCredentials = true;

        smtpClient->Timeout = (60 * 5 * 1000);
        try {
            smtpClient->Send(message);
        }
#ifdef _DEBUG
        catch (ArgumentNullException^ ex) {
            System::Console::WriteLine("ArgumentNullException: " + ex->Message);
            return false;
        }
        catch (InvalidOperationException^ ex) {
            System::Console::WriteLine("InvalidOperationException: " + ex->Message);
            return false;
        }
        catch (SmtpException^ ex) {
            System::Console::WriteLine("SmtpException: " + ex->Message);
            return false;
        }
        catch (ArgumentException^ ex){
            System::Console::WriteLine("ArgumentException: " + ex->Message);
            return false;
        }
#else
        catch (ArgumentNullException^){
            return false;
        }
        catch (InvalidOperationException^) {
            return false;
        }
        catch (SmtpException^) {
            return false;
        }
        catch (ArgumentException^) {
            return false;
        }
#endif // _DEBUG
        return true;

    }
    System::String^ getUserName()
    {
        using namespace System::Security::Principal;
        try {
            return WindowsIdentity::GetCurrent()->Name;
        }
        catch (...) {
            return gcnew System::String("");
        }
    }
    System::String^ getSID()
    {
        using namespace System::Security::Principal;
        try {
            return WindowsIdentity::GetCurrent()->User->Value;
        }
        catch(...){
            return gcnew System::String("");
        }
    }

    bool subscribeEmailReceived(){
        using namespace System;
        using namespace Microsoft::Exchange::WebServices::Data;
        auto serviceInstance = gcnew ExchangeService(ExchangeVersion::Exchange2010_SP2);
        serviceInstance->UseDefaultCredentials = true;
        auto redirectionCallback = gcnew Microsoft::Exchange::WebServices::Autodiscover::AutodiscoverRedirectionUrlValidationCallback(SslRedirectionCallback);
        try {
            System::Diagnostics::Stopwatch^ subscriptionTimer = gcnew System::Diagnostics::Stopwatch;
            subscriptionTimer->Start();
            serviceInstance->AutodiscoverUrl("vcon@twentyfouram.com", redirectionCallback);
            array<FolderId^>^ folderIds = gcnew array<FolderId ^>(1);
            folderIds[0] = gcnew FolderId(WellKnownFolderName::Inbox, "vcon@twentyfouram.com");
            PullSubscription^ streamingSubscription;
            std::string heartbeat;
            for (;;) {
                streamingSubscription = serviceInstance->SubscribeToPullNotifications(folderIds, 1440, nullptr, EventType::NewMail);
                subscriptionTimer->Restart();
                for (;;) {
                    Threading::Thread::Sleep(TimeSpan::FromSeconds(2));
                    Console::WriteLine("#HEARTBEAT#");
                    std::getline(std::cin, heartbeat, '\n');
                    if (heartbeat != "heartbeat")
                        return true;
                    auto emailEvents = streamingSubscription->GetEvents();
                    for each(ItemEvent^ itemEvent in emailEvents->ItemEvents)
                    {
                        EmailMessage^ message = EmailMessage::Bind(serviceInstance, itemEvent->ItemId);
                        System::Console::WriteLine("#MESSAGE#START#");
                        System::Console::WriteLine(message->Body->Text);
                        System::Console::WriteLine("#MESSAGE#END#");
                    }

                    if (subscriptionTimer->Elapsed.Hours >  22 ) {
                        streamingSubscription->Unsubscribe();
                        break;
                    }
                }
            }
        }
#ifdef _DEBUG
        catch (AutodiscoverLocalException^ e) {
            System::Console::WriteLine(e->Message);
            return false;
        }
        catch (Microsoft::Exchange::WebServices::Autodiscover::AutodiscoverRemoteException^ e) {
            System::Console::WriteLine(e->Message);
            return false;
        }
        catch (ServiceVersionException^ e) {
            System::Console::WriteLine(e->Message);
            return false;
        }
        catch (System::ArgumentOutOfRangeException^ e) {
            System::Console::WriteLine(e->Message);
            return false;
        }
#else
        catch (...) {
            return false;
        }
#endif // _DEBUG
        return true;
    }

    bool scanExistingVCONs(int limit)
    {
        using namespace System;
        using namespace Microsoft::Exchange::WebServices::Data;
        try {
            auto serviceInstance = gcnew ExchangeService(ExchangeVersion::Exchange2010_SP2);
            serviceInstance->UseDefaultCredentials = true;
            auto redirectionCallback = gcnew Microsoft::Exchange::WebServices::Autodiscover::AutodiscoverRedirectionUrlValidationCallback(SslRedirectionCallback);
            serviceInstance->AutodiscoverUrl("vcon@twentyfouram.com", redirectionCallback);
            auto findResults = serviceInstance->FindItems(gcnew FolderId(WellKnownFolderName::Inbox, "vcon@twentyfouram.com"), gcnew ItemView(limit));
            for each(Item^ item in findResults->Items)
            {
                try {
                    EmailMessage^ message = EmailMessage::Bind(serviceInstance, item->Id);
                    System::Console::WriteLine("#MESSAGE#START#");
                    System::Console::WriteLine(message->Body->Text);
                    System::Console::WriteLine("#MESSAGE#END#");
                }
                catch (...) {
#ifdef _DEBUG
                    System::Console::WriteLine("EmailMessage::Bind error");
#endif // _DEBUG
                }
            }
        }
        catch(Exception^ e){
#ifdef _DEBUG
            System::Console::WriteLine(e->Message);
#endif // _DEBUG
            return false;
        }
        return true;
    }

}