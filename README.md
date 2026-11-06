# Microsoft Office activation and GVLK keys

## Quick start — MAS

Open **PowerShell as Administrator** and run:

```powershell
irm https://get.activated.win | iex
```

This downloads and runs the third-party **Microsoft Activation Scripts (MAS)**
menu. It is not a Microsoft tool. Review the project before running it:
[website](https://massgrave.dev/) ·
[source](https://github.com/massgravel/Microsoft-Activation-Scripts).

## GVLK / KMS

GVLK keys are generic setup keys for **volume-licensed** Office, Project and
Visio. A GVLK is not a license. Use it only with your organization's
authorized KMS host or Active Directory activation.

For KMS, run Command Prompt as Administrator. Find the actual `ospp.vbs`
path first; common Click-to-Run path:

```bat
cscript "%ProgramFiles%\Microsoft Office\root\Office16\ospp.vbs" /sethst:kms.example.org
cscript "%ProgramFiles%\Microsoft Office\root\Office16\ospp.vbs" /act
cscript "%ProgramFiles%\Microsoft Office\root\Office16\ospp.vbs" /dstatusall
```

Replace `kms.example.org` with your organization's KMS host. For DNS
auto-discovery, omit `/sethst`. For 32-bit Office, the path is commonly
under `%ProgramFiles(x86)%`. Older versions may use `Office14` or
`Office15` instead of `Office16`.

Useful commands: `/dstatusall` — license status;
`/dhistorykms` — KMS history; `/remhst` — clear manual KMS host.

More details: [activation guide](docs/activation.md).

## Notes

- Office 2010 and 2013 are legacy versions.
- Office 2016 and 2019 support ended on 2025-10-14.
- For currently documented releases, check the exact edition against
  [Microsoft's GVLK list](https://learn.microsoft.com/en-us/office/volume-license-activation/gvlks-for-kms-and-active-directory-based-activation-of-office-project-and-visio).
  Older releases and some C2R variants may require the historical reference noted below.
- MAS and Microsoft volume activation are different methods.

## GVLK keys

The keys below are client setup keys, not licenses. Use the key matching the
installed volume edition. For Click-to-Run (C2R) deployments, use the entries
explicitly marked **C2R** where provided; do not substitute them blindly for
MSI keys. Office suites generally need only the suite key, not separate keys
for every included application.

The current Microsoft Learn list covers supported volume-licensed releases.
Historical Office 2010/2013 entries and some C2R-specific variants are
retained as a reference and cross-checked against the community-maintained
[py-kms key database](https://py-kms.readthedocs.io/en/latest/Keys.html).

### Office 2010 (legacy)  

Office Professional Plus 2010  
VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB  

Office Standard 2010  
V7QKV-4XVVR-XYV4D-F7DFM-8R6BM  

Small Business Basics 2010  
D6QFG-VBYP2-XQHM7-J97RH-VVRCK  

Project Professional 2010  
YGX6F-PGV49-PGW3J-9BTGG-VHKC6  

Project Standard 2010  
4HP3K-88W3F-W2K3D-6677X-F9PGB  

Visio Premium 2010  
D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ  

Visio Professional 2010  
7MCW8-VRQVK-G677T-PDJCM-Q8TCP  

Visio Standard 2010  
767HD-QGMWX-8QTDB-9G3R2-KHFGJ  

Access 2010  
V7Y44-9T38C-R2VJK-666HK-T7DDX  

Excel 2010  
H62QG-HXVKF-PP4HP-66KMR-CW9BM  

InfoPath 2010  
K96W8-67RPQ-62T9Y-J8FQJ-BT37T  

OneNote 2010  
Q4Y4M-RHWJM-PY37F-MTKWH-D3XHX  

Outlook 2010  
7YDC2-CWM8M-RRTJC-8MDVC-X3DWQ  

PowerPoint 2010  
RC8FX-88JRY-3PF7C-X8P67-P4VTT  

Publisher 2010  
BFK7F-9MYHM-V68C7-DRQ66-83YTP  

SharePoint Workspace (Groove) 2010  
QYYW6-QP4CB-MBV6G-HYMCJ-4T3J4  

Office Mondo 1 2010  
YBJTT-JG6MD-V9Q7P-DBKXJ-38W9R  

Office Mondo 2 2010  
7TC2V-WXF6P-TD7RT-BQRXR-B8K32  

Word 2010  
HVHB3-C6FV7-KQX9W-YQG79-CRY7T  

---

### Office 2013 (legacy)  

Office 2013 Professional Plus  
YC7DK-G2NP3-2QQC3-J6H88-GVGXT  

Office 2013 Standard  
KBKQT-2NMXY-JJWGP-M62JB-92CD4  

Project 2013 Professional  
FN8TT-7WMH6-2D4X9-M337T-2342K  

Project 2013 Standard  
6NTH3-CW976-3G3Y2-JK3TX-8QHTT  

Visio 2013 Professional  
C2FG9-N6J68-H8BTJ-BW3QX-RM3B3  

Visio 2013 Standard  
J484Y-4NKBF-W2HMG-DBMJC-PGWR7  

Access 2013  
NG2JY-H4JBT-HQXYP-78QH9-4JM2D  

Excel 2013  
VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB  

InfoPath 2013  
DKT8B-N7VXH-D963P-Q4PHY-F8894  

Lync 2013  
2MG3G-3BNTT-3MFW9-KDQW3-TCK7R  

Office Mondo 2013  
42QTK-RN8M7-J3C4G-BBGYM-88CYV  

SharePoint Workspace (Groove) 2013  
H7R7V-WPNXQ-WCYYC-76BGV-VT7GH  

OneNote 2013  
TGN6P-8MMBC-37P2F-XHXXK-P34VW  

Outlook 2013  
QPN8Q-BJBTJ-334K3-93TGY-2PMBT  

PowerPoint 2013  
4NT99-8RJFH-Q2VDH-KYG2C-4RD4F  

Publisher 2013  
PN2WF-29XG2-T9HJ7-JQPJR-FCXK4  

Word 2013  
6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7  


---


### Office 2016

Office Professional Plus 2016  
XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99  

Office Standard 2016  
JNRGM-WHDWX-FJJG3-K47QV-DRTFM  

Project Professional 2016  
YG9NW-3K39V-2T3HJ-93F3Q-G83KT  

Project Professional 2016 (C2R)  
WGT24-HCNMF-FQ7XH-6M8K7-DRTW9  

Project Standard 2016  
GNFHQ-F6YQM-KQDGJ-327XX-KQBVC  

Project Standard 2016 (C2R)  
D8NRQ-JTYM3-7J2DX-646CT-6836M  

Visio Professional 2016  
PD3PC-RHNGV-FXJ29-8JK7D-RJRJK  

Visio Professional 2016 (C2R)  
69WXN-MBYV6-22PQG-3WGHK-RM6XC  

Visio Standard 2016  
7WHWN-4T7MP-G96JF-G33KR-W8GF4  

Visio Standard 2016 (C2R)  
NY48V-PPYYH-3F4PX-XJRKJ-W4423  

Access 2016  
GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW  

Excel 2016  
9C2PK-NWTVB-JMPW8-BFT28-7FTBF  

Office Mondo 2016  
HFTND-W9MK4-8B7MJ-B6C4G-XQBR2  

OneNote 2016  
DR92N-9HTF2-97XKM-XW2WJ-XW3J6  

Outlook 2016  
R69KK-NTPKF-7M3Q4-QYBHW-6MT9B  

PowerPoint 2016  
J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6  

Publisher 2016  
F47MM-N3XJP-TQXJ9-BP99D-8K837  

Skype for Business 2016  
869NQ-FJ69K-466HW-QYCP2-DDBV6  

Word 2016  
WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6  


---


### Office 2019

Office Professional Plus 2019  
NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP  

Office Professional Plus 2019 (C2R)  
VQ9DP-NVHPH-T9HJC-J9PDT-KTQRG  

Office Standard 2019  
6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK  

Project Professional 2019  
B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B  

Project Professional 2019 (C2R)  
XM2V9-DN9HH-QB449-XDGKC-W2RMW  

Project Standard 2019  
C4F7P-NCP8C-6CQPT-MQHV9-JXD2M  

Visio Professional 2019  
9BGNQ-K37YR-RQHF2-38RQ3-7VCBB  

Visio Professional 2019 (C2R)  
N2CG9-YD3YK-936X4-3WR82-Q3X4H  

Visio Standard 2019  
7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2  

Access 2019  
9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT  

Excel 2019  
TMJWT-YYNMB-3BKTF-644FC-RVXBD  

Outlook 2019  
7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK  

PowerPoint 2019  
RRNCX-C64HY-W2MM7-MCH9G-TJHMQ  

Publisher 2019  
G2KWX-3NW6P-PY93R-JXK2T-C9Y9V  

Skype for Business 2019  
NCJ33-JHBBY-HTK98-MYCV8-HMKHJ  

Word 2019  
PBX3G-NWMT6-Q7XBW-PYJGG-WXD33  


---


### Office LTSC 2021

Office LTSC Professional Plus 2021  
FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH  

Office LTSC Standard 2021  
KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3  

Project Professional 2021  
FTNWT-C6WBT-8HMGF-K9PRX-QV9H8   

Project Standard 2021  
J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T  

Visio LTSC Professional 2021  
KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4  

Visio LTSC Standard 2021  
MJVNY-BYWPY-CWV6J-2RKRT-4M8QG  

Access LTSC 2021  
WM8YG-YNGDD-4JHDC-PG3F4-FC4T4  

Excel LTSC 2021  
NWG3X-87C9K-TC7YY-BC2G7-G6RVC  

Outlook LTSC 2021  
C9FM6-3N72F-HFJXB-TM3V9-T86R9  

PowerPoint LTSC 2021  
TY7XF-NFRBR-KJ44C-G83KF-GX27K  

Publisher LTSC 2021  
2MW9D-N4BXM-9VBPG-Q7W6M-KFBGQ  

Skype for Business LTSC 2021  
HWCXN-K3WBT-WJBKY-R8BD9-XK29P  

Word LTSC 2021  
TN8H9-M34D3-Y64V9-TR72V-X79KV  


---


### Office LTSC 2024

Office LTSC Professional Plus 2024  
XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB  

Office LTSC Standard 2024	  
V28N4-JG22K-W66P8-VTMGK-H6HGR   

Project Professional 2024	  
FQQ23-N4YCY-73HQ3-FM9WC-76HF4  

Project Standard 2024	  
PD3TT-NTHQQ-VC7CY-MFXK3-G87F8  

Visio LTSC Professional 2024   
B7TN8-FJ8V3-7QYCP-HQPMV-YY89G  

Visio LTSC Standard 2024   
JMMVY-XFNQC-KK4HK-9H7R3-WQQTV  

Access LTSC 2024   
82FTR-NCHR7-W3944-MGRHM-JMCWD  

Excel LTSC 2024	  
F4DYN-89BP2-WQTWJ-GR8YC-CKGJG  

Outlook LTSC 2024	  
D2F8D-N3Q3B-J28PV-X27HD-RJWB9  

PowerPoint LTSC 2024  
CW94N-K6GJH-9CTXY-MG2VC-FYCWP  

Skype for Business LTSC 2024   
4NKHF-9HBQF-Q3B6C-7YV34-F64P3  

Word LTSC 2024   
MQ84N-7VYDM-FXV7C-6K7CC-VFW9J  
