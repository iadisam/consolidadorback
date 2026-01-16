# API Consolidador Excel - FastAPI

Backend FastAPI para consolidar archivos Excel basado en una plantilla maestra.

## 🚀 Instalación

```bash
pip install -r requirements.txt
```

## ▶️ Ejecutar el servidor

```bash
python main.py
```

O con uvicorn:

```bash
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

La API estará disponible en: `http://localhost:8000`

Documentación interactiva: `http://localhost:8000/docs`

---

## 📋 Endpoints de la API

### **TEMPLATE** - Gestión de Plantillas

#### 1. GET `/api/template/download`
Descarga la plantilla actual cargada en el sistema.

**Response:** Archivo Excel (.xlsx o .xlsm)

**Ejemplo:**
```bash
curl -O http://localhost:8000/api/template/download
```

```javascript
// React/JavaScript
const response = await fetch('http://localhost:8000/api/template/download');
const blob = await response.blob();
const url = window.URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = 'plantilla.xlsm';
a.click();
```

---

#### 2. POST `/api/template/upload`
Sube la plantilla maestra y retorna las hojas disponibles.

**Request:**
- `template`: File (multipart/form-data)

**Response:**
```json
{
  "template_id": "a1b2c3d4e5f6",
  "template_name": "plantilla.xlsm",
  "sheet_names": ["Hoja1", "Hoja2", "Hoja3", "Resumen"]
}
```

**Ejemplo curl:**
```bash
curl -X POST http://localhost:8000/api/template/upload \
  -F "template=@plantilla.xlsm"
```

**Ejemplo JavaScript:**
```javascript
const formData = new FormData();
formData.append('template', templateFile);

const response = await fetch('http://localhost:8000/api/template/upload', {
  method: 'POST',
  body: formData
});

const data = await response.json();
console.log('Template ID:', data.template_id);
console.log('Hojas disponibles:', data.sheet_names);
```

---

### **CONSOLIDATE** - Proceso de Consolidación

#### 3. POST `/api/consolidate/upload`
Sube los archivos que serán consolidados.

**Request:**
- `files`: List[File] (multipart/form-data)

**Response:**
```json
{
  "session_id": "f6e5d4c3b2a1",
  "files_count": 5,
  "files": ["archivo1.xlsx", "archivo2.xlsx", "archivo3.xlsx", "archivo4.xlsx", "archivo5.xlsx"],
  "message": "Archivos subidos correctamente. Use POST /api/consolidate/process para iniciar la consolidación."
}
```

**Ejemplo curl:**
```bash
curl -X POST http://localhost:8000/api/consolidate/upload \
  -F "files=@archivo1.xlsx" \
  -F "files=@archivo2.xlsx" \
  -F "files=@archivo3.xlsx"
```

**Ejemplo JavaScript:**
```javascript
const formData = new FormData();
excelFiles.forEach(file => {
  formData.append('files', file);
});

const response = await fetch('http://localhost:8000/api/consolidate/upload', {
  method: 'POST',
  body: formData
});

const data = await response.json();
const sessionId = data.session_id;
console.log(`${data.files_count} archivos subidos`);
```

---

#### 4. POST `/api/consolidate/process`
Inicia el proceso de consolidación.

**Request:**
- `session_id`: String (form-data) - ID de sesión de los archivos subidos
- `excluded_sheets`: String (form-data, opcional) - Hojas a excluir separadas por coma

**Response:**
```json
{
  "task_id": "1a2b3c4d5e6f7g8h",
  "message": "Consolidación iniciada",
  "status_url": "/api/consolidate/status/1a2b3c4d5e6f7g8h",
  "included_sheets": ["Hoja1", "Hoja2", "Hoja3"],
  "excluded_sheets": ["Resumen"]
}
```

**Ejemplo curl:**
```bash
curl -X POST http://localhost:8000/api/consolidate/process \
  -F "session_id=f6e5d4c3b2a1" \
  -F "excluded_sheets=Resumen,Graficos"
```

**Ejemplo JavaScript:**
```javascript
const formData = new FormData();
formData.append('session_id', sessionId);
formData.append('excluded_sheets', 'Resumen,Graficos');

const response = await fetch('http://localhost:8000/api/consolidate/process', {
  method: 'POST',
  body: formData
});

const data = await response.json();
const taskId = data.task_id;
console.log('Task ID:', taskId);
```

---

#### 5. GET `/api/consolidate/status/{taskId}`
Obtiene el estado actual de una tarea de consolidación.

**Response:**
```json
{
  "task_id": "1a2b3c4d5e6f7g8h",
  "status": "processing",
  "progress": 45,
  "current_file": "archivo2.xlsx",
  "status_message": "Procesando archivo 2 de 5",
  "result_id": null,
  "error": null
}
```

Cuando está completo:
```json
{
  "task_id": "1a2b3c4d5e6f7g8h",
  "status": "completed",
  "progress": 100,
  "current_file": "Completado",
  "status_message": "Archivo consolidado listo para descargar",
  "result_id": "abc123def456",
  "error": null
}
```

**Ejemplo curl:**
```bash
curl http://localhost:8000/api/consolidate/status/1a2b3c4d5e6f7g8h
```

**Ejemplo JavaScript (Polling):**
```javascript
const pollStatus = async (taskId) => {
  const response = await fetch(`http://localhost:8000/api/consolidate/status/${taskId}`);
  const data = await response.json();
  
  console.log(`Progreso: ${data.progress}%`);
  console.log(`Estado: ${data.status_message}`);
  
  if (data.status === 'completed') {
    console.log('¡Consolidación completa!');
    console.log('Result ID:', data.result_id);
    return data.result_id;
  }
  
  if (data.status === 'error') {
    console.error('Error:', data.error);
    throw new Error(data.error);
  }
  
  return null;
};

// Polling cada segundo
const interval = setInterval(async () => {
  const resultId = await pollStatus(taskId);
  if (resultId) {
    clearInterval(interval);
    // Descargar archivo
    downloadFile(resultId);
  }
}, 1000);
```

---

#### 6. GET `/api/consolidate/download/{resultId}`
Descarga el archivo consolidado.

**Response:** Archivo Excel consolidado

**Ejemplo curl:**
```bash
curl -O http://localhost:8000/api/consolidate/download/abc123def456
```

**Ejemplo JavaScript:**
```javascript
const downloadFile = async (resultId) => {
  const url = `http://localhost:8000/api/consolidate/download/${resultId}`;
  
  // Opción 1: Abrir en nueva ventana
  window.open(url, '_blank');
  
  // Opción 2: Descargar programáticamente
  const response = await fetch(url);
  const blob = await response.blob();
  const downloadUrl = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = downloadUrl;
  a.download = 'REM_Consolidado.xlsm';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.URL.revokeObjectURL(downloadUrl);
};
```

---

## 🔄 Flujo Completo de Uso

### Paso a Paso

```javascript
// 1️⃣ SUBIR PLANTILLA
const uploadTemplate = async (file) => {
  const formData = new FormData();
  formData.append('template', file);
  
  const response = await fetch('http://localhost:8000/api/template/upload', {
    method: 'POST',
    body: formData
  });
  
  const data = await response.json();
  return {
    templateId: data.template_id,
    sheetNames: data.sheet_names
  };
};

// 2️⃣ SUBIR ARCHIVOS A CONSOLIDAR
const uploadFiles = async (files) => {
  const formData = new FormData();
  files.forEach(file => formData.append('files', file));
  
  const response = await fetch('http://localhost:8000/api/consolidate/upload', {
    method: 'POST',
    body: formData
  });
  
  const data = await response.json();
  return data.session_id;
};

// 3️⃣ INICIAR CONSOLIDACIÓN
const startConsolidation = async (sessionId, excludedSheets = []) => {
  const formData = new FormData();
  formData.append('session_id', sessionId);
  if (excludedSheets.length > 0) {
    formData.append('excluded_sheets', excludedSheets.join(','));
  }
  
  const response = await fetch('http://localhost:8000/api/consolidate/process', {
    method: 'POST',
    body: formData
  });
  
  const data = await response.json();
  return data.task_id;
};

// 4️⃣ MONITOREAR PROGRESO
const checkProgress = async (taskId) => {
  const response = await fetch(`http://localhost:8000/api/consolidate/status/${taskId}`);
  return await response.json();
};

// 5️⃣ DESCARGAR RESULTADO
const downloadResult = (resultId) => {
  window.open(`http://localhost:8000/api/consolidate/download/${resultId}`, '_blank');
};

// 🎯 FLUJO COMPLETO
const consolidateExcel = async (templateFile, excelFiles, excludedSheets) => {
  try {
    // 1. Subir plantilla
    console.log('📄 Subiendo plantilla...');
    const { sheetNames } = await uploadTemplate(templateFile);
    console.log('✅ Plantilla subida. Hojas:', sheetNames);
    
    // 2. Subir archivos
    console.log('📤 Subiendo archivos a consolidar...');
    const sessionId = await uploadFiles(excelFiles);
    console.log('✅ Archivos subidos. Session ID:', sessionId);
    
    // 3. Iniciar consolidación
    console.log('⚙️ Iniciando consolidación...');
    const taskId = await startConsolidation(sessionId, excludedSheets);
    console.log('✅ Consolidación iniciada. Task ID:', taskId);
    
    // 4. Monitorear progreso
    console.log('👀 Monitoreando progreso...');
    const pollInterval = setInterval(async () => {
      const status = await checkProgress(taskId);
      console.log(`📊 Progreso: ${status.progress}% - ${status.status_message}`);
      
      if (status.status === 'completed') {
        clearInterval(pollInterval);
        console.log('✅ ¡Consolidación completada!');
        
        // 5. Descargar resultado
        console.log('💾 Descargando archivo...');
        downloadResult(status.result_id);
      }
      
      if (status.status === 'error') {
        clearInterval(pollInterval);
        console.error('❌ Error:', status.error);
      }
    }, 1000);
    
  } catch (error) {
    console.error('❌ Error en el proceso:', error);
  }
};
```

---

## 🎨 Ejemplo de Integración con React

```jsx
import React, { useState } from 'react';

function ConsolidadorExcel() {
  const [templateFile, setTemplateFile] = useState(null);
  const [excelFiles, setExcelFiles] = useState([]);
  const [sheetNames, setSheetNames] = useState([]);
  const [excludedSheets, setExcludedSheets] = useState([]);
  const [progress, setProgress] = useState(0);
  const [status, setStatus] = useState('');
  const [taskId, setTaskId] = useState(null);

  const handleTemplateUpload = async (e) => {
    const file = e.target.files[0];
    setTemplateFile(file);

    const formData = new FormData();
    formData.append('template', file);

    const response = await fetch('http://localhost:8000/api/template/upload', {
      method: 'POST',
      body: formData
    });

    const data = await response.json();
    setSheetNames(data.sheet_names);
  };

  const handleFilesUpload = (e) => {
    setExcelFiles(Array.from(e.target.files));
  };

  const handleConsolidate = async () => {
    // 1. Subir archivos
    const formData = new FormData();
    excelFiles.forEach(file => formData.append('files', file));

    const uploadResponse = await fetch('http://localhost:8000/api/consolidate/upload', {
      method: 'POST',
      body: formData
    });

    const { session_id } = await uploadResponse.json();

    // 2. Iniciar consolidación
    const processFormData = new FormData();
    processFormData.append('session_id', session_id);
    if (excludedSheets.length > 0) {
      processFormData.append('excluded_sheets', excludedSheets.join(','));
    }

    const processResponse = await fetch('http://localhost:8000/api/consolidate/process', {
      method: 'POST',
      body: processFormData
    });

    const { task_id } = await processResponse.json();
    setTaskId(task_id);

    // 3. Monitorear progreso
    const interval = setInterval(async () => {
      const statusResponse = await fetch(`http://localhost:8000/api/consolidate/status/${task_id}`);
      const statusData = await statusResponse.json();

      setProgress(statusData.progress);
      setStatus(statusData.status_message);

      if (statusData.status === 'completed') {
        clearInterval(interval);
        window.open(`http://localhost:8000/api/consolidate/download/${statusData.result_id}`, '_blank');
      }
    }, 1000);
  };

  return (
    <div>
      <h1>Consolidador Excel</h1>
      
      <div>
        <h3>1. Subir Plantilla</h3>
        <input type="file" accept=".xlsx,.xlsm" onChange={handleTemplateUpload} />
      </div>

      {sheetNames.length > 0 && (
        <div>
          <h3>2. Seleccionar Hojas a Excluir</h3>
          {sheetNames.map(sheet => (
            <label key={sheet}>
              <input
                type="checkbox"
                checked={excludedSheets.includes(sheet)}
                onChange={(e) => {
                  if (e.target.checked) {
                    setExcludedSheets([...excludedSheets, sheet]);
                  } else {
                    setExcludedSheets(excludedSheets.filter(s => s !== sheet));
                  }
                }}
              />
              {sheet}
            </label>
          ))}
        </div>
      )}

      <div>
        <h3>3. Subir Archivos a Consolidar</h3>
        <input type="file" multiple accept=".xlsx,.xlsm" onChange={handleFilesUpload} />
      </div>

      <button onClick={handleConsolidate} disabled={!templateFile || excelFiles.length === 0}>
        Consolidar
      </button>

      {taskId && (
        <div>
          <h3>Progreso</h3>
          <progress value={progress} max="100" />
          <p>{progress}% - {status}</p>
        </div>
      )}
    </div>
  );
}

export default ConsolidadorExcel;
```

---

## 📊 Estados de las Tareas

| Estado | Descripción |
|--------|-------------|
| `processing` | La consolidación está en progreso |
| `completed` | La consolidación finalizó exitosamente |
| `error` | Ocurrió un error durante la consolidación |

---

## 🎯 Características

✅ **Mantiene macros VBA** - Soporta archivos .xlsm  
✅ **Suma inteligente** - Solo valores numéricos (no fórmulas)  
✅ **Preserva fórmulas** - Las fórmulas de la plantilla permanecen intactas  
✅ **Progreso en tiempo real** - Monitoreo vía polling con task_id  
✅ **Exclusión de hojas** - Especifica qué hojas no procesar  
✅ **Procesamiento asíncrono** - BackgroundTasks de FastAPI  
✅ **Sistema de sesiones** - Manejo seguro de múltiples usuarios  
✅ **IDs únicos** - task_id y result_id para rastreo preciso  

---

## 📁 Estructura de Carpetas

```
.
├── main.py              # API FastAPI
├── requirements.txt     # Dependencias
├── README.md           # Documentación
├── Dockerfile          # Configuración Docker
├── docker-compose.yml  # Orquestación
├── uploads/            # Archivos temporales subidos
├── templates/          # Plantillas maestras
└── results/            # Archivos consolidados finales
```

---

## ⚙️ Endpoints Adicionales

### GET `/api/health`
Verifica el estado de la API

**Response:**
```json
{
  "status": "healthy",
  "template_loaded": true,
  "template_name": "plantilla.xlsm",
  "active_sessions": 3,
  "active_tasks": 1
}
```

### DELETE `/api/cleanup`
Limpia archivos antiguos y sesiones completadas

### DELETE `/api/reset`
Reinicia completamente el estado de la aplicación

---

## 🐳 Docker

```bash
docker-compose up -d
```

---

## 🔒 Notas de Producción

- Usar Redis para estado compartido
- Implementar autenticación (JWT)
- Rate limiting
- Almacenamiento en S3
- Logging estructurado
- Monitoring con Prometheus

---

## 📝 Licencia

Uso interno - Todos los derechos reservados
