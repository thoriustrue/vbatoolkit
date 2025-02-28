import { Box, Button, Typography, CircularProgress, Alert, Accordion, AccordionSummary, AccordionDetails } from '@mui/material';
import ExpandMoreIcon from '@mui/icons-material/ExpandMore';
import ProcessLogs from './ProcessLogs';

// ... existing code ...

      {/* Logs */}
      {logs.length > 0 && (
        <ProcessLogs logs={logs} processType="VBA Code Extraction" />
      )}

// ... existing code ... 