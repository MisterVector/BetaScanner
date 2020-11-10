    Features, Tips, and Whats New.
    ==============================


	VERSION 1.0.3 (unofficial update by Vector)

    --> Updated the verification IP in BETAScannerINI.INI to the latest useast one

    --> Updated the verification IPs in VerificationIPs.txt

    --> The amount of proxies able to be tested is now 2,147,483,647

    --> Verification method for proxies has changed, as the old method no longer worked


	VERSION 1.0.2

    --> BetaScanner now supports HTTP scanning. While scanning on ports
    80, 8080, 3124, 3125, 3126, 3127, 3128 the proxy is assumed as http.
    
    --> You may also specify range types by placing @s4  or  @http at the end
    of your ranges or scanlist entries.


	VERSION 1.0.1

    --> The tray icon is updated with Verified, and granted count every 50 seconds.
    Upon scanning finish the icon will change from blue to green.

    --> Scan delay/thread accuracy tested at 98% consistancy; with much more efficiency than TLS.

    --> Auto-Completes can be edited, created, and removed by editing VerificationIPs.txt.
    Note that you cannot verify bNet proxies against non bNet IPs.

    --> Pressing return/enter at any editable field saves settings.

    --> Make sure your start / end IPs are setup properly as there is no error checking for this.
    If ports are not included in scanlist and ranges 1080 is assumed.

    --> Any text after the proxy in a scanlist is cutoff, to allow easy copying and pasting from
    websites.

    --> IMPORTANT: To those that assume the excessive denied proxy use is a program bug it IS NOT.
    A lot of proxies will decline connection request depending on the destination IP. Thus a
    lot of battleNet server IPs are blocked by a lot of proxies. Try changing the
    verification'destination' IP to attempt to fix this problem. 

