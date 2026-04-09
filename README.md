# PaperStream-AutomationScripts
For system admins to deploy Fujitsu Paperstream silently in their environment


To Download Fujitsu Setup Files:
https://www.pfu.ricoh.com/global/scanners/fi/dl/win-11-sp-11xxn.html

Once you've obtained the installers you need to run them once, and while at main install screen, copy the directory that is extracted and do this for all Paperstream setup files obtained from their website. In the extracted directories you should be able to locate the files needed. Because these utilize a wrapper, any other methods will fail to pass silent install parameters to the internal MSI.