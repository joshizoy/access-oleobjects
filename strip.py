def extract_real_file(ole_blob, out_path):
  signatures = {        
    b"%PDF": ".pdf",        
    b"PK\x03\x04": ".docx",
    b"\xD0\xCF\x11\xE0":".doc",  # old MS Office binary        
    b"\x89PNG": ".png"
  }    

  # Find file signature
  for sig, ext in signatures.items():
    idx = ole_blob.find(sig)
    if idx != -1:            
      with open (out_path + ext, "wb") as f:
        f.write(ole_blob[idx:])            
      print(f"Extracted {out_path}{ext}")            
      return    

  print("Unknown format")

# Usage after reading the binary
with open("Doc_1.bin", "rb") as f:
  blob = f.read()
extract_real_file(blob, "C:/ExportedFiles/Doc_1")
