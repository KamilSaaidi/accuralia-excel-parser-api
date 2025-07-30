# Excel/CSV Parser API

API FastAPI pour parser les fichiers Excel (XLSX, XLS) et CSV et extraire leur contenu.

## 🚀 Fonctionnalités

- **Support Excel** : XLSX (Office 2007+) et XLS (Office 97-2003)
- **Support CSV** : Fichiers CSV avec encodage UTF-8
- **Extraction complète** : Contenu de toutes les feuilles Excel
- **Données structurées** : Headers, lignes et métadonnées
- **API REST** : Endpoints simples à utiliser

## 📋 Endpoints

### GET /
Health check de l'API

### POST /parse-excel
Parse un fichier Excel/CSV

**Paramètres :**
- `fileData` : Contenu base64 du fichier
- `filename` : Nom du fichier avec extension

**Réponse :**
```json
{
  "success": true,
  "text": "Contenu extrait en texte",
  "method": "EXCEL_PARSING_XLSX",
  "confidence": 0.9,
  "filename": "document.xlsx",
  "file_type": "xlsx",
  "sheets_count": 2,
  "sheets_data": {
    "Sheet1": {
      "headers": ["Colonne1", "Colonne2"],
      "rows": [["Valeur1", "Valeur2"]],
      "shape": [1, 2]
    }
  },
  "textLength": 150
}
```

## 🛠️ Déploiement sur Railway

1. **Connectez votre repo GitHub à Railway**
2. **Railway détectera automatiquement le Dockerfile**
3. **L'API sera disponible sur** : `https://[votre-app].up.railway.app`

## 🔧 Intégration avec Supabase

Dans votre fonction `process-extraction`, remplacez l'appel Excel par :

```typescript
const excelApiUrl = 'https://[votre-app].up.railway.app/parse-excel';
const formData = new FormData();
formData.append('fileData', base64Data);
formData.append('filename', filename);

const response = await fetch(excelApiUrl, {
  method: 'POST',
  body: formData
});
```

## 📦 Dépendances

- **pandas** : Lecture des fichiers Excel/CSV
- **openpyxl** : Support XLSX
- **xlrd** : Support XLS
- **fastapi** : Framework web
- **uvicorn** : Serveur ASGI 