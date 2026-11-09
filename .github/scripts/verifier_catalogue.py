"""Contrôle du catalogue téléchargé depuis SharePoint avant publication.

Usage : python verifier_catalogue.py <fichier_telecharge> <fichier_en_ligne> <code_http>

Écrit dans $GITHUB_OUTPUT :
  decision = publier | identique | bloque
  resume   = une ligne pour le message de commit
et produit rapport.md (repris dans le résumé du job et, en cas de blocage, dans le ticket d'alerte).

Bloquant (rien n'est publié, la version en ligne reste) :
  téléchargement en échec, page web au lieu du fichier, fichier illisible,
  onglet Formations ou colonne Titre introuvable, aucune formation publiée,
  nombre de formations publiées divisé par deux ou plus d'un coup.
Non bloquant (signalé dans le rapport) : tout le reste.
"""
import datetime as dt
import os
import re
import sys
import unicodedata

from openpyxl import load_workbook


def key(s):
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn").lower()
    return re.sub(r"[^a-z0-9]", "", s)


def find_sheet(wb, *names):
    wanted = {key(n) for n in names}
    for ws in wb.worksheets:
        if key(ws.title) in wanted:
            return ws
    return None


def rows_of(ws):
    return [list(r) for r in ws.iter_rows(values_only=True)]


def col(header, *aliases):
    keys = [key(h) for h in header]
    for a in aliases:
        if a in keys:
            return keys.index(a)
    for i, k in enumerate(keys):
        if k and any(k.startswith(a) for a in aliases):
            return i
    return None


def val(row, i):
    if i is None or i >= len(row) or row[i] is None:
        return ""
    return str(row[i]).strip()


def published(statut):
    s = key(statut)
    return s == "" or s.startswith("pub") or s.startswith("oui")


def parse_date(v):
    if isinstance(v, (dt.date, dt.datetime)):
        return True
    s = str(v or "").strip()
    if not s:
        return None
    return bool(re.match(r"^\d{4}-\d{1,2}-\d{1,2}", s) or re.match(r"^\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4}$", s))


class Bloque(Exception):
    pass


def analyser(path):
    """Retourne (stats, avertissements). Lève Bloque si le fichier ne peut pas être publié."""
    try:
        wb = load_workbook(path, data_only=True, read_only=False)
    except Exception as e:  # fichier corrompu ou format inattendu
        raise Bloque(f"Le fichier ne s'ouvre pas comme un classeur Excel ({type(e).__name__}).")

    ws = find_sheet(wb, "Formations", "Catalogue")
    if ws is None:  # même tolérance que la page : un onglet dont la première ligne contient « Titre »
        ws = next((w for w in wb.worksheets if any(key(c.value) == "titre" for c in next(w.iter_rows(max_row=1), []))), None)
    if ws is None:
        raise Bloque("L'onglet « Formations » est introuvable. A-t-il été renommé ou supprimé ?")
    rows = rows_of(ws)
    if not rows:
        raise Bloque("L'onglet « Formations » est vide.")
    head = rows[0]
    c_titre = col(head, "titre", "intitule")
    if c_titre is None:
        raise Bloque("La colonne « Titre » est introuvable dans l'onglet Formations.")
    c_id, c_statut = col(head, "id", "code"), col(head, "statut")
    c_dom, c_duree = col(head, "domaine", "thematique"), col(head, "duree")

    warn = []
    listes = find_sheet(wb, "Listes")
    domaines = set()
    if listes is not None:
        lrows = rows_of(listes)
        cd = col(lrows[0], "domaines", "domaine", "thematique") if lrows else None
        if cd is not None:
            for r in lrows[1:]:
                v = val(r, cd)
                if not v:
                    break
                domaines.add(key(v))

    ids, nb_pub, nb_total = {}, 0, 0
    for n, r in enumerate(rows[1:], start=2):
        titre = val(r, c_titre)
        if not titre:
            continue
        nb_total += 1
        ident = val(r, c_id).upper().replace(" ", "")
        if published(val(r, c_statut)):
            nb_pub += 1
        where = f"Formations, ligne {n} ({titre[:40]})"
        if not ident:
            warn.append(f"{where} : pas d'ID.")
        elif ident in ids:
            warn.append(f"{where} : ID {ident} déjà utilisé ligne {ids[ident]}.")
        else:
            ids[ident] = n
        dom = val(r, c_dom)
        if not dom:
            warn.append(f"{where} : thématique vide.")
        elif domaines and key(dom) not in domaines:
            warn.append(f"{where} : thématique « {dom} » absente de l'onglet Listes.")
        duree = val(r, c_duree)
        if duree and not re.search(r"\d+(?:[.,]\d+)?\s*h", duree, re.I):
            warn.append(f"{where} : durée « {duree} » sans nombre d'heures.")

    if nb_pub == 0:
        raise Bloque("Aucune formation n'est au statut « Publié » : le site serait vide.")

    nb_sess = 0
    ss = find_sheet(wb, "Sessions", "Calendrier")
    if ss is not None:
        srows = rows_of(ss)
        if srows:
            h = srows[0]
            cf, cs = col(h, "formationid", "formation", "id"), col(h, "datededebut", "debut", "date")
            for n, r in enumerate(srows[1:], start=2):
                fid = val(r, cf).upper().replace(" ", "")
                if not fid or len(fid) > 20:
                    continue
                nb_sess += 1
                if fid not in ids:
                    warn.append(f"Sessions, ligne {n} : formation « {fid} » introuvable dans l'onglet Formations.")
                ok = parse_date(r[cs] if cs is not None and cs < len(r) else None)
                if not ok:
                    warn.append(f"Sessions, ligne {n} : date de début manquante ou illisible.")

    reg = find_sheet(wb, "Réglages", "Reglages", "Parametres")
    if reg is not None:
        R = {key(a): val([b], 0) for a, b, *_ in (list(r) + [None, None] for r in reg.iter_rows(min_row=2, values_only=True)) if a}
        if not R.get("emailinscriptions"):
            warn.append("Réglages : « E-mail inscriptions » est vide, les demandes d'inscription partiront sans destinataire.")
        if key(R.get("datesfictivesdemonstration", "")).startswith("oui"):
            warn.append("Réglages : « Dates fictives (démonstration) » vaut encore Oui.")

    return {"publiees": nb_pub, "total": nb_total, "sessions": nb_sess}, warn


def main():
    nouveau, en_ligne, code = sys.argv[1], sys.argv[2], sys.argv[3]
    hote = (sys.argv[4] if len(sys.argv) > 4 else "").lower()
    out = os.environ.get("GITHUB_OUTPUT")
    lignes = []

    def sortie(decision, resume=""):
        with open("rapport.md", "w", encoding="utf-8") as f:
            f.write("\n".join(lignes) + "\n")
        if out:
            with open(out, "a", encoding="utf-8") as f:
                f.write(f"decision={decision}\nresume={resume}\n")
        print(f"decision={decision}")
        print("\n".join(lignes))

    def bloquer(msg):
        lignes.extend(["### Synchronisation bloquée", "", f"**{msg}**", "",
                       "La version en ligne n'a pas changé. Corrigez le fichier sur SharePoint : "
                       "le robot réessaie à chaque passage (plusieurs fois par jour, ou tout de suite avec Run workflow) et ce ticket se ferme tout seul quand tout repart."])
        sortie("bloque")

    if "login.microsoftonline" in hote or "login.live" in hote:
        return bloquer("Le lien demande une connexion Amicial. Dans « Paramètres du lien », choisissez "
                       "« Toute personne disposant du lien », cliquez sur Appliquer, puis copiez le lien à nouveau.")
    if code in ("401", "403"):
        return bloquer(f"SharePoint refuse le téléchargement (code {code}). Causes les plus fréquentes : le lien n'est "
                       "pas de type « Toute personne disposant du lien » (« Personnes dans AMICIAL » ne marche pas), "
                       "ou l'option « Bloquer le téléchargement » est cochée dans les paramètres du lien.")
    if code == "404":
        return bloquer("Lien introuvable (code 404) : le lien a été supprimé, ou le fichier supprimé ou déplacé vers un autre site.")
    if code != "200":
        return bloquer(f"Téléchargement impossible (code {code}). Le lien de partage a peut-être expiré, "
                       "été supprimé, ou le fichier a été déplacé.")
    with open(nouveau, "rb") as f:
        debut = f.read(4)
    if debut[:2] != b"PK":
        return bloquer("Le lien renvoie une page web au lieu du fichier. Vérifiez que le lien est de type "
                       "« Toute personne disposant du lien » (lecture seule).")

    if os.path.exists(en_ligne):
        with open(nouveau, "rb") as a, open(en_ligne, "rb") as b:
            if a.read() == b.read():
                lignes.append("Aucun changement : le fichier SharePoint est identique à la version en ligne.")
                return sortie("identique")

    try:
        stats, warn = analyser(nouveau)
    except Bloque as e:
        return bloquer(str(e))

    if os.path.exists(en_ligne):
        try:
            avant, _ = analyser(en_ligne)
            if avant["publiees"] >= 10 and stats["publiees"] < avant["publiees"] / 2:
                return bloquer(f"Le nombre de formations publiées passerait de {avant['publiees']} à "
                               f"{stats['publiees']}. Publication suspendue par sécurité. Si c'est voulu, "
                               "déposez le fichier à la main dans le dépôt (Add file → Upload files).")
        except Bloque:
            pass

    resume = f"{stats['publiees']} formations publiées, {stats['sessions']} sessions"
    lignes.extend(["### Catalogue publié", "", f"{resume} (sur {stats['total']} lignes).", ""])
    if warn:
        lignes.append(f"**{len(warn)} point(s) à vérifier** (non bloquants) :")
        lignes.append("")
        lignes.extend(f"- {w}" for w in warn[:60])
        if len(warn) > 60:
            lignes.append(f"- … et {len(warn) - 60} autre(s).")
    else:
        lignes.append("Aucun oubli détecté.")
    sortie("publier", resume)


if __name__ == "__main__":
    main()
