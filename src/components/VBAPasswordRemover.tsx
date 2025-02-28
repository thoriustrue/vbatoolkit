import { Box, Button, Typography, CircularProgress, Alert } from '@mui/material';
import ProcessLogs from './ProcessLogs';

const VBAPasswordRemover = () => {
  // ... existing code ...

  return (
    <Box>
      <Typography variant="h5" gutterBottom>
        VBA Password Remover
      </Typography>
      <Typography variant="body1" paragraph>
        Upload an Excel file with VBA password protection to remove the password and unlock the VBA project.
      </Typography>
      
      {/* File upload section */}
      <FileUploader
        onFileSelected={handleFileSelected}
        acceptedFileTypes={['.xlsm', '.xlsb', '.xla', '.xlam']}
        maxFileSize={50 * 1024 * 1024} // 50MB
        disabled={processing}
      />
      
      {/* Processing status */}
      {processing && (
        <Box display="flex" alignItems="center" mt={2}>
          <CircularProgress size={24} sx={{ mr: 2 }} />
          <Typography>
            {progressText} ({Math.round(progress * 100)}%)
          </Typography>
        </Box>
      )}
      
      {/* Success message */}
      {unprotectedFile && (
        <Alert severity="success" sx={{ mt: 2, mb: 2 }}>
          VBA password removal completed successfully!
          <Button 
            variant="outlined" 
            size="small" 
            onClick={handleDownload}
            sx={{ ml: 2 }}
          >
            Download Unprotected File
          </Button>
        </Alert>
      )}
      
      {/* Error message */}
      {error && (
        <Alert severity="error" sx={{ mt: 2 }}>
          {error}
        </Alert>
      )}
      
      {/* Logs */}
      {logs.length > 0 && (
        <ProcessLogs logs={logs} processType="VBA Password Removal" />
      )}
    </Box>
  );
};

export default VBAPasswordRemover; 