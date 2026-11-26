import csv
from pathlib import Path
import re
import tkinter as tk
from tkinter import filedialog

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font
except ImportError:
    print("Manca 'openpyxl'. Installalo dal terminale con: py -m pip install openpyxl")
    raise SystemExit(1)


# --- IMPOSTAZIONI TELEFONI E FILTRI ---
# Formatta i numeri USA a 10 cifre come (813) 810-0005 (o +1 (813) 810-0005 per 11 cifre con 1 iniziale)
FORMATTA_US_10_CIFRE = True
# Se True, scarta numeri composti da una sola cifra ripetuta (es. 1111111111, 0000000000)
ESCLUDI_NUMERI_FINTI = True
# Se True, scarta contatti il cui nome contiene la parola "test"
ESCLUDI_CONTATTI_TEST = True


def scegli_file_contatti() -> list[Path]:
    """Apre la finestra di Windows per selezionare uno o più file Excel (.xlsx) o CSV (.csv)."""
    root = tk.Tk()
    root.withdraw()
    percorsi = filedialog.askopenfilenames(
        title="Seleziona uno o più file Excel o CSV dei contatti",
        filetypes=[
            ("File Contatti (Excel e CSV)", "*.xlsx *.csv"),
            ("File Excel", "*.xlsx"),
            ("File CSV", "*.csv"),
        ],
    )
    root.destroy()
    return [Path(p) for p in percorsi]


def pulisci_testo(valore) -> str:
    if valore is None:
        return ""
    return " ".join(str(valore).strip().split()).title()


def unisci_nome_cognome(first_name, last_name) -> str:
    first = pulisci_testo(first_name)
    last = pulisci_testo(last_name)
    return f"{first} {last}".strip()


def pulisci_telefono(valore) -> str:
    """Pulisce il numero e formatta i numeri USA a 10 o 11 cifre."""
    if valore is None:
        return ""
    s = str(valore).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]

    cifre = re.sub(r"\D", "", s)
    if not cifre:
        return ""

    if FORMATTA_US_10_CIFRE:
        if len(cifre) == 10:
            return f"({cifre[:3]}) {cifre[3:6]}-{cifre[6:]}"
        if len(cifre) == 11 and cifre.startswith("1"):
            return f"+1 ({cifre[1:4]}) {cifre[4:7]}-{cifre[7:]}"

    return cifre


def pulisci_email(valore) -> str:
    if valore is None:
        return ""
    return str(valore).strip().lower()


def e_numero_finto(solo_cifre: str) -> bool:
    if len(solo_cifre) < 7:
        return True
    return len(set(solo_cifre)) == 1


def leggi_righe_tabella(path: Path) -> list[list]:
    """Legge le righe sia da un file .xlsx sia da un file .csv."""
    estensione = path.suffix.lower()

    if estensione == ".xlsx":
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        return [list(riga) for riga in ws.iter_rows(values_only=True)]

    if estensione == ".csv":
        for codifica in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                with path.open("r", encoding=codifica, newline="") as f:
                    campione = f.read(4096)
                    f.seek(0)
                    try:
                        dialetto = csv.Sniffer().sniff(campione, delimiters=",;\t|")
                    except csv.Error:
                        dialetto = csv.excel
                    lettore = csv.reader(f, dialetto)
                    return [riga for riga in lettore]
            except UnicodeDecodeError:
                continue
        raise ValueError(f"Impossibile decodificare il file CSV: {path.name}")

    raise ValueError(f"Formato file non supportato: {path.suffix}")


def carica_contatti_file(path: Path) -> tuple[list[dict], int, list[str]]:
    """Legge e pulisce un singolo file (.xlsx o .csv) restituendo contatti e dettagli scarti."""
    righe = leggi_righe_tabella(path)
    if not righe:
        return [], 0, []

    intestazioni = [str(h).strip().lower() if h is not None else "" for h in righe[0]]

    contatti = []
    senza_telefono = 0
    esclusi_dettaglio = []

    for idx, riga in enumerate(righe[1:], start=2):
        record = dict(zip(intestazioni, riga))

        nome_completo = unisci_nome_cognome(
            record.get("first name") or record.get("nome"),
            record.get("last name") or record.get("cognome"),
        )
        telefono = pulisci_telefono(record.get("phone") or record.get("telefono"))
        email = pulisci_email(record.get("email"))
        gender = pulisci_testo(record.get("gender"))

        if not nome_completo and not telefono and not email and not gender:
            continue

        if not telefono:
            senza_telefono += 1
            continue

        solo_cifre = re.sub(r"\D", "", telefono)
        if ESCLUDI_NUMERI_FINTI and e_numero_finto(solo_cifre):
            esclusi_dettaglio.append(
                f"[{path.name}] Riga {idx} (numero finto/corto): {nome_completo or 'Senza nome'} ({telefono})"
            )
            continue

        if ESCLUDI_CONTATTI_TEST and re.search(r"\btest\w*\b", nome_completo, re.IGNORECASE):
            esclusi_dettaglio.append(
                f"[{path.name}] Riga {idx} (contatto test): {nome_completo} ({telefono})"
            )
            continue

        contatti.append({
            "nome": nome_completo,
            "telefono": telefono,
            "telefoni": [telefono],
            "email": email,
            "gender": gender,
            "origine": path.name,
        })

    return contatti, senza_telefono, esclusi_dettaglio


def chiave_persona(contatto: dict) -> str:
    """Crea una chiave per riconoscere la stessa persona (email se presente, altrimenti nome)."""
    if contatto["email"]:
        return f"email:{contatto['email']}"
    nome_norm = " ".join(contatto["nome"].lower().split())
    return f"nome:{nome_norm}" if nome_norm else ""


def unisci_e_pulisci_file(file_list: list[Path]) -> tuple[list[dict], dict]:
    contatti_uniti = []
    duplicati = []
    nomi_diversi_stesso_tel = []
    numeri_uniti_stessa_persona = []
    esclusi_dettaglio = []
    senza_telefono = 0

    numeri_visti = {}        # chiave_tel -> contatto già salvato
    persone_viste = {}       # chiave_persona -> contatto già salvato

    for percorso in file_list:
        contatti_file, s_tel, filt_dettaglio = carica_contatti_file(percorso)
        senza_telefono += s_tel
        esclusi_dettaglio.extend(filt_dettaglio)

        for c in contatti_file:
            chiave_tel = re.sub(r"\D", "", c["telefono"])
            chiave_pers = chiave_persona(c)

            # 1. Il numero è già presente
            if chiave_tel in numeri_visti:
                esistente = numeri_visti[chiave_tel]
                nome_esistente = (esistente["nome"] or "").strip().lower()
                nome_nuovo = (c["nome"] or "").strip().lower()

                if nome_esistente and nome_nuovo and nome_esistente != nome_nuovo:
                    nomi_diversi_stesso_tel.append(
                        f"[{percorso.name}] Tel {c['telefono']}: '{c['nome']}' vs già salvato '{esistente['nome']}'"
                    )
                else:
                    duplicati.append(f"[{percorso.name}] {c['nome']} ({c['telefono']})")

                # Se mancava email o gender nel contatto esistente, completiamo i campi vuoti
                if not esistente["email"] and c["email"]:
                    esistente["email"] = c["email"]
                if not esistente["gender"] and c["gender"]:
                    esistente["gender"] = c["gender"]
                continue

            # 2. Numero nuovo, ma stessa persona (stessa email o stesso nome) già presente
            if chiave_pers and chiave_pers in persone_viste:
                esistente = persone_viste[chiave_pers]
                esistente["telefoni"].append(c["telefono"])
                esistente["telefono"] = " / ".join(esistente["telefoni"])
                numeri_visti[chiave_tel] = esistente

                if not esistente["email"] and c["email"]:
                    esistente["email"] = c["email"]
                if not esistente["gender"] and c["gender"]:
                    esistente["gender"] = c["gender"]

                numeri_uniti_stessa_persona.append(
                    f"[{percorso.name}] {esistente['nome'] or c['nome']}: aggiunto tel {c['telefono']}"
                )
                continue

            # 3. Nuovo contatto valido
            numeri_visti[chiave_tel] = c
            if chiave_pers:
                persone_viste[chiave_pers] = c
            contatti_uniti.append(c)

    return contatti_uniti, {
        "senza_telefono": senza_telefono,
        "esclusi_dettaglio": esclusi_dettaglio,
        "duplicati": duplicati,
        "nomi_diversi_stesso_tel": nomi_diversi_stesso_tel,
        "numeri_uniti_stessa_persona": numeri_uniti_stessa_persona,
    }


def salva_excel_pulito(contatti: list[dict], out_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Contatti_Puliti"

    ws.append(["Nome", "Telefono", "Email", "Gender"])
    for cella in ws[1]:
        cella.font = Font(bold=True)

    for c in sorted(contatti, key=lambda x: x["nome"]):
        ws.append([c["nome"], c["telefono"], c["email"], c["gender"]])

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4

    wb.save(out_path)


def salva_vcf_rubrica(contatti: list[dict], vcf_path: Path) -> None:
    """Crea un file .vcf (vCard 3.0) importabile in rubrica."""
    righe_vcf = []
    for c in sorted(contatti, key=lambda x: x["nome"]):
        nome = c["nome"].replace("\n", " ").strip() or "Senza Nome"
        parti = nome.split(" ", 1)
        first = parti[0]
        last = parti[1] if len(parti) > 1 else ""

        scheda = [
            "BEGIN:VCARD",
            "VERSION:3.0",
            f"FN:{nome}",
            f"N:{last};{first};;;",
        ]
        for tel in c.get("telefoni") or [c["telefono"]]:
            if tel:
                scheda.append(f"TEL;TYPE=CELL:{tel}")
        if c["email"]:
            scheda.append(f"EMAIL;TYPE=INTERNET:{c['email']}")
        scheda.append("END:VCARD")
        righe_vcf.extend(scheda)

    vcf_path.write_text("\n".join(righe_vcf) + "\n", encoding="utf-8")


def salva_report_testo(
    report_path: Path,
    num_file: int,
    totale_validi: int,
    stats: dict,
) -> None:
    """Salva un report dettagliato con totali, duplicati, unioni e voci escluse."""
    esclusi = stats["esclusi_dettaglio"]
    duplicati = stats["duplicati"]
    nomi_diversi = stats["nomi_diversi_stesso_tel"]
    numeri_uniti = stats["numeri_uniti_stessa_persona"]

    righe = [
        "RIEPILOGO PULIZIA CONTATTI",
        "=" * 42,
        f"File elaborati: {num_file}",
        f"Contatti unici salvati: {totale_validi}",
        f"Righe saltate senza telefono: {stats['senza_telefono']}",
        f"Righe escluse (numeri finti / test): {len(esclusi)}",
        f"Duplicati identici rimossi: {len(duplicati)}",
        f"Stesso numero ma nome diverso (segnalati): {len(nomi_diversi)}",
        f"Stessa persona con numero aggiunto: {len(numeri_uniti)}",
        "",
    ]

    if esclusi:
        righe.append("DETTAGLIO VOCI ESCLUSE (NUMERI FINTI / TEST):")
        righe.extend(f" - {item}" for item in esclusi)
        righe.append("")

    if numeri_uniti:
        righe.append("NUMERI AGGIUNTI ALLA STESSA PERSONA:")
        righe.extend(f" - {item}" for item in numeri_uniti)
        righe.append("")

    if nomi_diversi:
        righe.append("SEGNALAZIONI - STESSO NUMERO CON NOME DIVERSO:")
        righe.extend(f" - {item}" for item in nomi_diversi)
        righe.append("")

    if duplicati:
        righe.append("DETTAGLIO DUPLICATI RIMOSSI:")
        righe.extend(f" - {item}" for item in duplicati)
        righe.append("")

    report_path.write_text("\n".join(righe), encoding="utf-8")


def main() -> None:
    file_selezionati = scegli_file_contatti()
    if not file_selezionati:
        print("Nessun file selezionato.")
        return

    contatti, stats = unisci_e_pulisci_file(file_selezionati)
    cartella_base = file_selezionati[0]
    excel_out = cartella_base.with_name("contatti_uniti_puliti.xlsx")
    vcf_out = cartella_base.with_name("contatti_uniti_puliti.vcf")
    report_out = cartella_base.with_name("report_pulizia.txt")

    salva_excel_pulito(contatti, excel_out)
    salva_vcf_rubrica(contatti, vcf_out)
    salva_report_testo(report_out, len(file_selezionati), len(contatti), stats)

    print(f"File elaborati: {len(file_selezionati)}")
    print(f"Fatto! Contatti unici salvati: {len(contatti)}")
    print(f" - Excel: {excel_out}")
    print(f" - Rubrica (.vcf): {vcf_out}")
    print(f" - Report: {report_out}")
    print(f"Righe saltate senza telefono: {stats['senza_telefono']}")
    print(f"Righe escluse (numeri finti / test): {len(stats['esclusi_dettaglio'])}")
    print(f"Stessa persona con numero aggiunto: {len(stats['numeri_uniti_stessa_persona'])}")
    print(f"Stesso numero ma nome diverso: {len(stats['nomi_diversi_stesso_tel'])}")
    print(f"Duplicati rimossi: {len(stats['duplicati'])} (dettagli completi in {report_out.name})")


if __name__ == "__main__":
    main()
