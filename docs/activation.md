# Office volume activation: client and KMS notes

This guide is for administrators deploying **volume-licensed** Office,
Project, or Visio in an organization. It does not provide a way to activate
retail, Microsoft 365 Apps, or otherwise unlicensed installations.

## Choose the activation method

| Method | Where it fits | Client requirements |
| --- | --- | --- |
| KMS | Organization-managed activation service on the network | Volume-licensed Office with its GVLK; access to the KMS host |
| Active Directory-based activation | Domain-joined computers in an AD DS environment | Volume-licensed Office with its GVLK; access to the AD activation object |
| MAK | One-time activation against Microsoft activation services | A MAK issued to the organization |

A GVLK is a generic client setup key. It is not the KMS host key, a MAK, or
proof of a license entitlement.

## KMS behavior

- KMS clients normally discover the host through the DNS SRV record
  `_vlmcs._tcp`. Manual host configuration should be the exception.
- The default KMS service port is TCP 1688.
- Office KMS has a client activation threshold of five unique client requests.
  The host does not activate Office clients until it has received the required
  count.
- A successful KMS activation is valid for 180 days. Clients periodically
  renew their activation while they can reach the KMS host.
- The KMS host must be prepared for Office with the matching Office Volume
  License Pack and an organization-issued KMS host key. A Windows KMS host
  configuration alone does not automatically enable Office activation.

## Find and run ospp.vbs

The script is installed with Office. Do not assume a single path for every
release, architecture, or installation technology.

Typical current Click-to-Run Office 2016/2019/LTSC paths:

- 64-bit Office: `%ProgramFiles%\Microsoft Office\root\Office16\ospp.vbs`
- 32-bit Office on 64-bit Windows:
  `%ProgramFiles(x86)%\Microsoft Office\root\Office16\ospp.vbs`

Older MSI installations may use `Office14` (Office 2010),
`Office15` (Office 2013), or a different Office16 path. Locate the actual
file before running commands.

From an elevated Command Prompt, use the correct path:

```bat
cscript "C:\path\to\ospp.vbs" /dstatusall
cscript "C:\path\to\ospp.vbs" /dhistorykms
```

Only if DNS discovery is not configured or cannot be used, an administrator
may explicitly assign the organization's KMS host:

```bat
cscript "C:\path\to\ospp.vbs" /sethst:kms.example.org
cscript "C:\path\to\ospp.vbs" /act
cscript "C:\path\to\ospp.vbs" /dstatusall
```

To clear the explicit host and return to automatic discovery:

```bat
cscript "C:\path\to\ospp.vbs" /remhst
```

Do not point clients at arbitrary public KMS servers. Use only the activation
infrastructure authorized by your organization's licensing agreement.

## Troubleshooting checklist

1. Confirm the installed SKU is a volume-licensed edition, not retail or
   Microsoft 365 Apps.
2. Check the license state with `/dstatusall`. Confirm the expected product
   and channel are present.
3. Verify DNS resolution of the KMS SRV record:
   `nslookup -type=srv _vlmcs._tcp`.
4. Verify network reachability to the discovered KMS host on TCP 1688 (or the
   port configured by the administrator).
5. Confirm the KMS host has the Office Volume License Pack installed, an
   activated Office KMS host key, and support for the Office release.
6. Check `/dhistorykms` and the activation error code returned by `/act`.
7. After a successful activation, close and reopen Office applications if
   their UI still displays stale licensing information.

## Official Microsoft documentation

- [Activate Office by using KMS](https://learn.microsoft.com/en-us/office/volume-license-activation/activate-office-by-using-kms)
- [Configure a KMS host for Office](https://learn.microsoft.com/en-us/office/volume-license-activation/configure-a-kms-host-computer-for-office)
- [Activate Office by using Active Directory](https://learn.microsoft.com/en-us/office/volume-license-activation/activate-office-by-using-active-directory)
- [Manage Office volume activation tools](https://learn.microsoft.com/en-us/office/volume-license-activation/tools-to-manage-volume-activation-of-office)
- [Troubleshoot Office volume activation](https://learn.microsoft.com/en-us/office/volume-license-activation/troubleshoot-volume-activation-of-office)
