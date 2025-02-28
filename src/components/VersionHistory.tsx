import React, { useEffect, useState } from 'react';
import { Box, Typography, List, ListItem, ListItemText, Divider, CircularProgress } from '@mui/material';

interface VersionHistoryProps {
  onClose?: () => void;
}

const VersionHistory: React.FC<VersionHistoryProps> = ({ onClose }) => {
  const [changelog, setChangelog] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const fetchChangelog = async () => {
      try {
        setLoading(true);
        // Fetch the CHANGELOG.md file from the repository
        const response = await fetch('https://raw.githubusercontent.com/yourusername/vbatoolkit/main/CHANGELOG.md');
        
        if (!response.ok) {
          throw new Error(`Failed to fetch changelog: ${response.status}`);
        }
        
        const text = await response.text();
        setChangelog(text);
        setError(null);
      } catch (err) {
        console.error('Error fetching changelog:', err);
        setError('Failed to load version history. Please try again later.');
      } finally {
        setLoading(false);
      }
    };

    fetchChangelog();
  }, []);

  // Parse the changelog into sections
  const parseChangelog = (text: string) => {
    if (!text) return [];
    
    // Split by version headers (## [x.x.x])
    const sections = text.split(/## \[\d+\.\d+\.\d+\]/);
    const versionHeaders = text.match(/## \[\d+\.\d+\.\d+\]/g) || [];
    
    // Combine headers with content
    return sections.slice(1).map((content, index) => ({
      version: versionHeaders[index]?.replace('## [', '').replace(']', '') || 'Unknown',
      content: content.trim()
    }));
  };

  const versions = parseChangelog(changelog);

  return (
    <Box sx={{ 
      width: '100%', 
      maxHeight: '80vh',
      overflowY: 'auto', // Add scrollbar
      padding: 2 
    }}>
      <Typography variant="h5" gutterBottom>
        Version History
      </Typography>
      
      {loading ? (
        <Box sx={{ display: 'flex', justifyContent: 'center', p: 4 }}>
          <CircularProgress />
        </Box>
      ) : error ? (
        <Typography color="error">{error}</Typography>
      ) : (
        <List>
          {versions.map((version, index) => (
            <React.Fragment key={index}>
              <ListItem alignItems="flex-start">
                <ListItemText
                  primary={`Version ${version.version}`}
                  secondary={
                    <Box component="div" sx={{ whiteSpace: 'pre-line', mt: 1 }}>
                      {version.content}
                    </Box>
                  }
                />
              </ListItem>
              {index < versions.length - 1 && <Divider component="li" />}
            </React.Fragment>
          ))}
        </List>
      )}
    </Box>
  );
};

export default VersionHistory; 