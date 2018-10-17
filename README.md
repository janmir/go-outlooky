# :rocket: Overview
go-outlooky automation tool for Outlook 2016*
> *Should work with older versions but have not tested it.

# :book: OLE Reference 
- [OLE Documentation](https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook?view=outlook-pia)
- [Outlook MailItem](https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook._mailitem?view=outlook-pia)
- [Code Samples](http://techsupt.winbatch.com/webcgi/webbatch.exe?techsupt/nftechsupt.web+WinBatch/OLE~COM~ADO~CDO~ADSI~LDAP/OLE~and~Outlook+OLE~and~OUTLOOK~read~mail~other~than~inbox.txt)
- [Default Folder](https://docs.microsoft.com/en-us/office/vba/api/outlook.namespace.getdefaultfolder)

# OLE Outlook Objects
- [Namespace Object](https://docs.microsoft.com/en-us/office/vba/api/outlook.namespace)
- [Namespace Classes](https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook?view=outlook-pia)

# Snippets
> Match Japanese Characters
```js
 [\u3000-\u303f\u3040-\u309f\u30a0-\u30ff\uff00-\uff9f\u4e00-\u9faf\u3400-\u4dbf]
  -------------_____________-------------_____________-------------_____________
   Punctuation   Hiragana     Katakana    Full-width       CJK      CJK Ext. A
                                            Roman/      (Common &      (Rare)    
                                          Half-width    Uncommon)
                                           Katakana

 一-龯|぀-ゖ|ァ-ヺ|ｦ-ﾝ|ㇰ-ㇿ|Ａ-Ｚ|ａ-ｚ|０-９（）ー［］ー・
```