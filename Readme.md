# Skype for Business and Let's Encrypt

I wanted to be able to use Let's Encrypt certificates for my Skype for Business lab Edge Server and could not find a blog or something else that explained how to do that with a script that I could schedule with the Task Scheduler. So I created my own script (there is always room for improvement) that would do that for me.

I already have a Let's Encrypt certificate in my lab for the Web Services and I use pfSense with the HaProxy package as a reverse proxy for that. pfSense also has a Let's Encrypt package that can automatically request the certificates that you need for Skype for Business.

## Requirements

If you want to use this script to retrieve and assign Let's Encrypt certificates on your Edge server you need do some preparations.
- Download the free version of the Mongoose webserver or search for your own standalone webserver. https://cesanta.com/binary.html
- Download the Let's Encrypt commandline client. https://github.com/do-know/Crypt-LE/releases

Place the mongoose-free.exe and le64.exe in the same directory as the script.

Open port 80 on your firewall for the three Edge Server IP Addresses. Let's Encrypt will use port 80 to check the request. The firewall on the Edge server does not have port 80 open so the script will open this temporary during the request fase and closes the port when the script is done.

## Script
Start the script and the rest will happen automatically.

Start with the 'test' Let's Encrypt server:
```powershell
cd <script location>
.\Update-Certificates.ps1 -PfxPassword <YourPassword> -SipFQDN sip.domain.net -WebFQDN web.domain.net -AvFQDN av.domain.net 
```

Open certlm.msc and check if you have three new certificates in you computer store. If this worked, switch to the 'live' Let's Encrypt server and get some real certificates that work for 3 months.

```powershell
.\Update-Certificates.ps1 -PfxPassword <YourPassword> -SipFQDN sip.domain.net -WebFQDN web.domain.net -AvFQDN av.domain.net -live
```

Again, open certlm.msc and check if you have three real certificates from Let's Encrypt. These three certificates should also be assigned to the Skype for Business Edge Roles. You can check this with the following command:
```powershell
Get-CsCertificate
```

If everything works as expected, you can schedule the script with the Windows Task Scheduler.
