from flask import Flask, render_template, request, redirect, url_for, session
import json
from datetime import datetime
from dateutil.relativedelta import relativedelta  
import pandas as pd
import io
from flask import send_file
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from flask import Response
from matplotlib.ticker import MaxNLocator

app = Flask(__name__)
app.secret_key = 'lumis_fibre_optique_mdp'

USERNAME = 'lumis.user'
PASSWORD = 'lumis@2025'



def charger_projets():
    try:
        with open('projets.json', 'r', encoding='utf-8') as f:
            projets = json.load(f)
            max_id = 0
            for i, projet in enumerate(projets):
                if 'id' not in projet:
                    projet['id'] = i + 1  
                max_id = max(max_id, projet['id'])

            sauvegarder_projets(projets)  
            return projets
    except FileNotFoundError:
        return []


def sauvegarder_projets(projets):
    with open('projets.json', 'w', encoding='utf-8') as f:
        json.dump(projets, f, indent=4, ensure_ascii=False)


# Listes fixes
CLIENTS = ["EURO FIBER", "SUDALYS", "AXIANS-NIMES", "AXIANS-GASQ", "SOGETREL", "PRIME SAS"]
TACHES = ["APS", "APD", "DOE", "CAPFT", "COMAC", "GC", "CA\\DFT", "DT", "NEXLOOP"]
ETATS = ["À faire", "En cours", "Faite", "Bloqué"]
REALISATEURS = ["Yassine", "Omar", "Mohammed", "Rajae", "Oumaima"]

DUREES_TACHES_HEURES = {
    "APS": 3,
    "APD": 2,
    "DOE": 3,
    "CAPFT": 2,
    "COMAC": 9,
    "GC": 4,
    "CA\\DFT": 2,
    "DT": 4,
    "NEXLOOP": 27
}


from functools import wraps

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function


@app.route("/")
@login_required
def index():
    projets = charger_projets()
    
    mois_str = request.args.get("mois")
    if mois_str:
        try:
            mois_courant = datetime.strptime(mois_str, "%Y-%m")
        except ValueError:
            mois_courant = datetime.now().replace(day=1)
    else:
        mois_courant = datetime.now().replace(day=1)

    session['mois_courant'] = mois_courant.strftime("%Y-%m")
    projets_mois = []
    for p in projets:
        try:
            date_rec = datetime.strptime(p["date_reception"], "%Y-%m-%d")
            if date_rec.year == mois_courant.year and date_rec.month == mois_courant.month:
                projets_mois.append(p)
        except Exception:
            continue


    mois_prec = (mois_courant - relativedelta(months=1)).strftime("%Y-%m")
    mois_suiv = (mois_courant + relativedelta(months=1)).strftime("%Y-%m")

    mois_courant_str = mois_courant.strftime("%B %Y").capitalize()
    mois_courant_param = mois_courant.strftime("%Y-%m")  

    return render_template(
        "index.html",
        projets=projets_mois,
        mois_courant_str=mois_courant_str,  
        mois_courant_param=mois_courant_param, 
        mois_prec=mois_prec,
        mois_suiv=mois_suiv
    )


@app.route("/export_excel")
@login_required
def export_excel():
    projets = charger_projets()

    mois_str = request.args.get("mois")
    if mois_str:
        try:
            mois_courant = datetime.strptime(mois_str, "%Y-%m")
        except ValueError:
            mois_courant = datetime.now().replace(day=1)
    else:
        mois_courant = datetime.now().replace(day=1)

    # Filtrer les projets du mois
    projets_mois = []
    for p in projets:
        try:
            date_rec = datetime.strptime(p["date_reception"], "%Y-%m-%d")
            if date_rec.year == mois_courant.year and date_rec.month == mois_courant.month:
                projets_mois.append(p)
        except:
            continue

    if not projets_mois:
        return "Aucun projet trouvé pour ce mois.", 404

    df = pd.DataFrame(projets_mois)

    # Générer le diagramme
    compte_pondere = {r: 0 for r in REALISATEURS}
    for p in projets_mois:
        if p.get("etat") != "Bloqué":
            realisateur = p.get("realisateur")
            tache = p.get("tache")
            if realisateur in compte_pondere:
                duree = DUREES_TACHES_HEURES.get(tache, 1)
                compte_pondere[realisateur] += duree

    fig = Figure(figsize=(10, 4))
    ax = fig.subplots()

    noms = list(compte_pondere.keys())
    valeurs = list(compte_pondere.values())
    couleurs = {
        "Yassine": "#df6ecd",
        "Omar": "#83c876",
        "Mohammed": "#a9afc1",
        "Rajae": "#ffe5a0",
        "Oumaima": "#ffcfc9"
    }
    bar_colors = [couleurs.get(nom, "#004080") for nom in noms]

    ax.bar(noms, valeurs, color=bar_colors)
    for i, val in enumerate(valeurs):
        texte = f"{val:.1f}".replace(".", ",")
        ax.text(i, val + 0.1, texte, ha='center', va='bottom')
    ax.set_title(f"Répartition du temps - {mois_courant.strftime('%B %Y').capitalize()}")
    ax.set_ylabel("Charge pondérée")
    ax.grid(axis='y', linestyle='--', alpha=0.6)

    output = io.BytesIO()
    fig.savefig(output, format="png", bbox_inches="tight")
    output.seek(0)
    img_data = output.read()

    # Créer le fichier Excel
    excel_output = io.BytesIO()
    with pd.ExcelWriter(excel_output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Projets")
        
        workbook  = writer.book
        worksheet = workbook.add_worksheet("Diagramme")
        writer.sheets["Diagramme"] = worksheet

        # Insérer l'image dans la feuille "Diagramme"
        worksheet.insert_image('B2', 'graphique.png', {'image_data': io.BytesIO(img_data)})

    excel_output.seek(0)

    nom_fichier = f"projets_{mois_courant.strftime('%Y_%m')}.xlsx"
    return send_file(
        excel_output,
        as_attachment=True,
        download_name=nom_fichier,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/ajouter", methods=["GET", "POST"])
@login_required
def ajouter():
    if request.method == "POST":
        projets = charger_projets()
        nouvel_id = max([p['id'] for p in projets], default=0) + 1
        nouveau_projet = {
            "id": nouvel_id,
            "client": request.form["client"],
            "tache": request.form["tache"],
            "projet": request.form["projet"],
            "date_reception": request.form["date_reception"],
            "realisateur": request.form["realisateur"],
            "etat": request.form["etat"],
            "date_envoi": request.form["date_envoi"]
        }
        projets.append(nouveau_projet)
        sauvegarder_projets(projets)
        try:
            date_rec = datetime.strptime(request.form["date_reception"], "%Y-%m-%d")
            mois_param = date_rec.strftime("%Y-%m")
            return redirect(url_for("index", mois=mois_param))
        except:
            return redirect(url_for("index"))

    return render_template("add_project.html", clients=CLIENTS, taches=TACHES, etats=ETATS, realisateurs=REALISATEURS)


@app.route("/modifier/<int:id>", methods=["GET", "POST"])
@login_required
def modifier(id):
    projets = charger_projets()
    projet = next((p for p in projets if p["id"] == id), None)

    if not projet:
        return redirect(url_for("index"))

    if request.method == "POST":
        projet["client"] = request.form["client"]
        projet["tache"] = request.form["tache"]
        projet["projet"] = request.form["projet"]
        projet["date_reception"] = request.form["date_reception"]
        projet["realisateur"] = request.form["realisateur"]
        projet["etat"] = request.form["etat"]
        projet["date_envoi"] = request.form["date_envoi"]

        sauvegarder_projets(projets)

        mois_param = session.get('mois_courant', None)
        if mois_param:
            return redirect(url_for("index", mois=mois_param))
        else:
            return redirect(url_for("index"))


    return render_template("modifier_project.html",
                           projet=projet,
                           clients=CLIENTS,
                           taches=TACHES,
                           etats=ETATS,
                           realisateurs=REALISATEURS)


@app.route("/supprimer/<int:id>")
@login_required
def supprimer(id):
    projets = charger_projets()
    projet_a_supprimer = next((p for p in projets if p["id"] == id), None)

    if projet_a_supprimer:
        projets = [p for p in projets if p["id"] != id]  
        sauvegarder_projets(projets)

        mois_param = session.get('mois_courant', None)
        if mois_param:
            return redirect(url_for("index", mois=mois_param))

        return redirect(url_for("index"))

    return redirect(url_for("index"))

from flask import request

from datetime import date

@app.route("/termine/<int:id>")
@login_required
def marquer_termine(id):
    projets = charger_projets()
    projet = next((p for p in projets if p["id"] == id), None)
    if projet:
        projet["date_envoi"] = date.today().strftime("%Y-%m-%d")
        projet["etat"] = "Faite"
        sauvegarder_projets(projets)
    mois_param = session.get('mois_courant', None)
    return redirect(url_for("index", mois=mois_param) if mois_param else url_for("index"))


@app.route("/graph_taches")
@login_required
def graph_taches():
    mois_str = request.args.get("mois")
    projets = charger_projets()
    
    if mois_str:
        try:
            mois_courant = datetime.strptime(mois_str, "%Y-%m")
        except ValueError:
            mois_courant = None
    else:
        mois_courant = None

    if mois_courant:
        projets = [
            p for p in projets
            if "date_reception" in p
            and datetime.strptime(p["date_reception"], "%Y-%m-%d").year == mois_courant.year
            and datetime.strptime(p["date_reception"], "%Y-%m-%d").month == mois_courant.month
            and p.get("etat") != "Bloqué"
        ]

    compte_pondere = {r: 0 for r in REALISATEURS}
    for p in projets:
        realisateur = p.get("realisateur")
        tache = p.get("tache")
        if realisateur in compte_pondere:
            duree = DUREES_TACHES_HEURES.get(tache, 1)
            compte_pondere[realisateur] += duree



    fig = Figure(figsize=(8, 6))
    ax = fig.subplots()
    plt.rcParams.update({'font.size': 10, 'font.family': 'DejaVu Sans', 'font.weight': 'light'})

    noms = list(compte_pondere.keys())
    valeurs = list(compte_pondere.values())


    couleurs = {
        "Yassine": "#df6ecd",
        "Omar": "#83c876",
        "Mohammed": "#a9afc1",
        "Rajae": "#ffe5a0",
        "Oumaima": "#ffcfc9"
    }
    bar_colors = [couleurs.get(nom, "#004080") for nom in noms]

    ax.bar(noms, valeurs, color=bar_colors)
    for i, val in enumerate(valeurs):
        texte = f"{val:.1f}".replace(".", ",")
        ax.text(i, val + 0.1, texte, ha='center', va='bottom')
    ax.set_xlabel("Réalisateurs", color="#6b7082")
    ax.set_ylabel("Temps passé par (h)", color="#6b7082")
    ax.set_title(f"Temps total passé par réalisateur - {mois_str if mois_str else 'Tous les mois'}")
    ax.grid(axis='y', linestyle='--', alpha=0.7)

    from matplotlib.ticker import MaxNLocator
    ax.yaxis.set_major_locator(MaxNLocator(integer=True))
    from matplotlib.ticker import FuncFormatter
    ax.yaxis.set_major_formatter(FuncFormatter(lambda x, _: f"{x:.1f}".replace('.', ',')))

    output = io.BytesIO()
    fig.savefig(output, format="png", bbox_inches='tight')
    output.seek(0)
    return Response(output.getvalue(), mimetype="image/png")

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]
        if username == USERNAME and password == PASSWORD:
            session["logged_in"] = True
            return redirect(url_for("index"))
        else:
            erreur = "Nom d'utilisateur ou mot de passe incorrect."
            return render_template("login.html", erreur=erreur)
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
