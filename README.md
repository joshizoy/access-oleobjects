# Automate Extracting Documents from MS Access OLE Object Columns
Extracting documents (like Word, Excel, PDF, or images) that are stored inside MS Access OLE Object fields can be tricky because Access does not store the raw file directly. Instead, it wraps the file in an OLE header that depends on the application used to insert the object (Word, Excel, Paint, etc.). That means you usually have to strip off the OLE wrapper before you can recover the original file.

Here's a detailed guide with options depending on the approach you want to take.

## Understanding OLE Storage in Access
When you insert a file into an OLE Object field, Access doesn’t just store the file. Instead, it stores:

- OLE Header – metadata about the embedding application and object type.
- Raw File Data – the actual content of the file (but wrapped).

So simply reading the binary field gives you a blob with extra bytes at the beginning.

## Using VBA inside Access
VBA (Visual Basic for Applications) is a programming language developed by Microsoft that’s built into Office applications like Excel, Access, and Word. Therefore, this is the first and most obvious choice to automate extracting documents from MS Access OLE Objects. Use script extract.bas from this repository for this purpose.

## Using .NET (C# or VB.NET)
VB.NET (Visual Basic .NET) is an object-oriented programming language developed by Microsoft. It’s part of the .NET framework and is used to build desktop, web, and mobile applications. Unlike VBA, which is limited to Office apps, VB.NET is a general-purpose language with modern features like inheritance, exception handling, and strong typing.

So, if you are extracting data programmatically outside MS Access, use script extract_vbntet.cls from this repository.

## Stripping the OLE Header
The hard part is getting the real file out. The header size varies depending on how the file was inserted:

- Bitmaps: OLE header usually 78 bytes.
- Word, Excel, PDF: often larger, unpredictable.

Sometimes MS Access stores the full file name at the start. In order to deal with it, look for the magic number (file signature) inside the binary blob. For example:

- PDF → %PDF (25 50 44 46)
- DOCX/ZIP → PK (50 4B 03 04)
- PNG → 89 50 4E 47

Once found, strip everything before that offset and save the rest as a new file. File strip.py contains Python example of doing this (after exporting raw OLE).

## Third-Party Tools
If you don't want to deal with OLE headers manually, consider using those tools for full automation of extracting documents from MS Access OLE Objects:

- Access OLE Export – commercial tool ($95) to extract images and files from OLE Object fields of MS Access database
- [Access-to-MySQL](https://www.convert-in.com/access-to-mysql) – commercial tool ($79) to migrate MS Access data including OLE Objects to MySQL, MariaDB or Percona (both on premises and cloud platforms)
- [Access-to-PostgreSQL](https://www.convert-in.com/acc2pgs) – commercial tool ($79) to migrate MS Access data including OLE Objects to PostgreSQL (both on premises and cloud platforms)
