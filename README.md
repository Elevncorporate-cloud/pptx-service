# Micro-service PPTX/PDF — La Machine à CV

Service Python qui reçoit le JSON envoyé par Lovable, ouvre le template
PowerPoint d'origine, remplace les jetons par les données utilisateur, puis
exporte un PDF identique au design Canva/PowerPoint.

## 1. Préparer le template

Édite `templates/cv-1.pptx` (déjà placé à partir de ton fichier `CV 1.pptx`)
et remplace les textes du modèle par les jetons suivants — sans toucher au
style ni à la position des zones :

| Champ                  | Jeton à coller dans le PPTX |
|------------------------|------------------------------|
| Nom complet            | `{{FULL_NAME}}`              |
| Métier                 | `{{JOB_TITLE}}`              |
| Email                  | `{{EMAIL}}`                  |
| Téléphone              | `{{PHONE}}`                  |
| Ville                  | `{{CITY}}`                   |
| Adresse                | `{{ADDRESS}}`                |
| Accroche               | `{{SUMMARY}}`                |
| Bloc expériences       | `{{EXPERIENCES}}`            |
| Bloc formations        | `{{EDUCATION}}`              |
| Compétences            | `{{SKILLS}}`                 |
| Soft skills            | `{{SOFT_SKILLS}}`            |
| Langues                | `{{LANGUAGES}}`              |
| Intérêts               | `{{INTERESTS}}`              |
| Photo                  | Une zone de texte contenant `{{PHOTO}}` (sera remplacée par l'image) |

Pour ajouter un nouveau modèle, dépose-le dans `templates/<slug>.pptx` et
appelle l'API avec `"templateId": "<slug>"`.

## 2. Lancer en local

```bash
pip install -r requirements.txt
# LibreOffice requis pour la conversion PDF
brew install libreoffice          # macOS
# ou : sudo apt install libreoffice  (Ubuntu)

export PPTX_API_KEY="choisis-une-cle"
export PUBLIC_BASE_URL="http://localhost:8080"
uvicorn main:app --reload --port 8080
```

## 3. Déployer (Render / Railway / Fly)

Le `Dockerfile` est prêt — installe `libreoffice` et lance `uvicorn`.

1. Crée un nouveau service Web depuis ce dossier.
2. Variables d'environnement :
   - `PPTX_API_KEY` : clé partagée avec Lovable
   - `PUBLIC_BASE_URL` : URL publique du service (ex. `https://cv.onrender.com`)
3. Build/Run géré par le Dockerfile (port `$PORT`).

## 4. Brancher Lovable

Dans Lovable, ajoute les deux secrets :

- `PPTX_API_URL` = `https://<ton-service>/generate`
- `PPTX_API_KEY` = même valeur que ci-dessus

Lovable POSTera automatiquement le JSON structuré (voir `src/lib/pdf/payload.ts`)
et recevra `{ pptxUrl, pdfUrl }`.

## 5. Test rapide

```bash
curl -X POST http://localhost:8080/generate \
  -H "Authorization: Bearer $PPTX_API_KEY" \
  -H "Content-Type: application/json" \
  -d '{"templateId":"cv-1","fullName":"Jean Dupont","jobTitle":"Commercial B2B","email":"jean@x.fr","phone":"0612345678","city":"Paris","summary":"5 ans en BtoB.","experiences":[],"education":[],"skills":["Vente","CRM"],"softSkills":[],"languages":[],"interests":[]}'
```
