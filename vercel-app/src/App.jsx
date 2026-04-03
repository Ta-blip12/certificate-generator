import React, { useState } from 'react';
import './App.css';

export default function App() {
  const [docxFile, setDocxFile] = useState(null);
  const [excelFile, setExcelFile] = useState(null);
  const [courseName, setCourseName] = useState('');
  const [issueDate, setIssueDate] = useState('');
  const [studyLoad, setStudyLoad] = useState('');
  const [startNumber, setStartNumber] = useState('');
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');
  const [messageType, setMessageType] = useState('info');
  const [progress, setProgress] = useState(0);

  const showMessage = (text, type = 'info') => {
    setMessage(text);
    setMessageType(type);
  };

  const clearMessage = () => {
    setMessage('');
  };

  const handleFileChange = (e, setter) => {
    const file = e.target.files[0];
    if (file) {
      setter(file);
    }
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    clearMessage();

    if (!docxFile) {
      showMessage('Selecteer het Word-certificaattemplate', 'error');
      return;
    }
    if (!excelFile) {
      showMessage('Selecteer de Excel-bestand met deelnemers', 'error');
      return;
    }
    if (!courseName) {
      showMessage('Vul de naam van de nascholing in', 'error');
      return;
    }
    if (!issueDate) {
      showMessage('Vul de datum van uitgifte in', 'error');
      return;
    }
    if (!studyLoad) {
      showMessage('Vul de studiebelasting in', 'error');
      return;
    }
    if (!startNumber) {
      showMessage('Vul het beginnummer in', 'error');
      return;
    }

    setLoading(true);
    setProgress(0);

    try {
      const formData = new FormData();
      formData.append('docx', docxFile);
      formData.append('excel', excelFile);
      formData.append('courseName', courseName);
      formData.append('issueDate', issueDate);
      formData.append('studyLoad', studyLoad);
      formData.append('startNumber', startNumber);

      // Simulate progress
      const progressInterval = setInterval(() => {
        setProgress((prev) => Math.min(prev + 10, 90));
      }, 200);

      const response = await fetch('/api/generate', {
        method: 'POST',
        body: formData,
      });

      clearInterval(progressInterval);

      if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error || 'Generation failed');
      }

      setProgress(100);

      // Download the ZIP file
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'certificaten.zip';
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      showMessage('✓ Certificaten gegenereerd! Download is gestart.', 'success');
      resetForm();

    } catch (error) {
      showMessage(`Fout: ${error.message}`, 'error');
    } finally {
      setLoading(false);
      setProgress(0);
    }
  };

  const resetForm = () => {
    setDocxFile(null);
    setExcelFile(null);
    setCourseName('');
    setIssueDate('');
    setStudyLoad('');
    setStartNumber('');
  };

  return (
    <div className="container">
      <div className="header">
        <h1>Certificaatgenerator</h1>
        <p>Maak in één keer alle certificaten voor je deelnemers</p>
      </div>

      {message && (
        <div className={`message message-${messageType}`}>
          {message}
        </div>
      )}

      <form onSubmit={handleSubmit}>
        <div className="section">
          <div className="section-title">Stap 1: Bestanden</div>

          <div className="form-group">
            <label>Word-certificaattemplate (.docx)</label>
            <div className="file-input-wrapper">
              <label className="file-label">Klik om bestand te uploaden</label>
              <input
                type="file"
                accept=".docx"
                onChange={(e) => handleFileChange(e, setDocxFile)}
              />
            </div>
            {docxFile && <div className="file-name">✓ {docxFile.name}</div>}
          </div>

          <div className="form-group">
            <label>Deelnemerslijst (Excel .xlsx)</label>
            <div className="file-input-wrapper">
              <label className="file-label">Klik om bestand te uploaden</label>
              <input
                type="file"
                accept=".xlsx"
                onChange={(e) => handleFileChange(e, setExcelFile)}
              />
            </div>
            {excelFile && <div className="file-name">✓ {excelFile.name}</div>}
          </div>
        </div>

        <div className="section">
          <div className="section-title">Stap 2: Certificaatgegevens</div>

          <div className="form-group">
            <label>Naam van de nascholing</label>
            <input
              type="text"
              value={courseName}
              onChange={(e) => setCourseName(e.target.value)}
              placeholder="bijv. Introduction to Nordoff-Robbins Music Therapy"
              disabled={loading}
            />
          </div>

          <div className="form-group">
            <label>Datum van uitgifte</label>
            <input
              type="text"
              value={issueDate}
              onChange={(e) => setIssueDate(e.target.value)}
              placeholder="bijv. 5 maart 2026"
              disabled={loading}
            />
          </div>

          <div className="form-group">
            <label>Studiebelasting (in uren)</label>
            <input
              type="text"
              value={studyLoad}
              onChange={(e) => setStudyLoad(e.target.value)}
              placeholder="bijv. 16"
              disabled={loading}
            />
          </div>

          <div className="form-group">
            <label>Beginnummer</label>
            <input
              type="text"
              value={startNumber}
              onChange={(e) => setStartNumber(e.target.value)}
              placeholder="bijv. 2604.01"
              disabled={loading}
            />
          </div>
        </div>

        <div className="button-group">
          <button
            type="button"
            onClick={resetForm}
            disabled={loading}
          >
            Wissen
          </button>
          <button
            type="submit"
            className="button-primary"
            disabled={loading}
          >
            {loading ? 'Bezig...' : 'Certificaten genereren ↗'}
          </button>
        </div>

        {loading && (
          <div className="progress">
            <div className="progress-bar">
              <div
                className="progress-fill"
                style={{ width: `${progress}%` }}
              ></div>
            </div>
            <p className="progress-text">
              {progress < 100 ? 'Certificaten worden gegenereerd...' : 'Download gestart...'}
            </p>
          </div>
        )}
      </form>
    </div>
  );
}
