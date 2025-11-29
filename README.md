![RAR](https://github.com/user-attachments/assets/fa554897-32ff-44b0-8e20-034e77b9bb58)  

# Integrarea Verificare-RAR-ITP pentru Home Assistant

Aceasta integrare verifica valabilitatea Inspectiei Tehnice Periodice pe baza Seriei VIN.

<img width="478" height="231" alt="Screenshot 2025-07-21 at 18 43 40" src="https://github.com/user-attachments/assets/cf21140a-9958-4a5d-a220-e71fd6257b04" />

Pentru ca aceasta integrare sa functioneze foloseste OCR.Space pentru a citi si trimite codul captha, verificare se face odata la 30 zile.

Pentru a obtine API Key pentru OCR.Space urmatorii pasi trebuiesc urmati:

🔹 Step 1: Go to OCR.Space API Portal
🔹 Step 2: Enter your email address
🔹 Step 3: Check your inbox and confirm the details
🔹 Step 4: Check your inbox for the free API key (subject: "Your OCR.Space API Key")

## De unde descarc fisierele

Ai doua optiuni pentru a instala integrarea:

1. **Prin HACS (recomandat)**
   - Deschide HACS → _Integrations_ → _Custom repositories_.
   - Adauga URL-ul acestui depozit GitHub (`https://github.com/mariusmotea/Verificare-RAR-ITP`) ca _Integration_ si apoi instaleaza pachetul.
   - Reporneste Home Assistant dupa instalare.

2. **Instalare manuala**
   - Descarca arhiva ZIP a ultimei versiuni din sectiunea **Releases** a acestui depozit GitHub: [https://github.com/mariusmotea/Verificare-RAR-ITP/releases](https://github.com/mariusmotea/Verificare-RAR-ITP/releases).
   - Dupa descarcare, extrage arhiva si copiaza folderul `custom_components/rar_itp_checker` in directorul `config/custom_components` din instanta ta Home Assistant (creeaza structura daca nu exista).
   - Reporneste Home Assistant pentru a incarca integrarea.


