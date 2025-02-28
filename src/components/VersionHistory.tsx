import React, { useEffect, useState } from 'react';
import { Box, Typography, List, ListItem, ListItemText, Divider, CircularProgress } from '@mui/material';
import CHANGELOG_DATA from '../data/changelog';

interface VersionHistoryProps {
  onClose?: () => void;
}

const VersionHistory: React.FC<VersionHistoryProps> = ({ onClose }) => {
  // Use the static CHANGELOG_DATA from the codebase
  // This avoids the need to fetch the CHANGELOG.md file
  
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
      
      <List>
        {CHANGELOG_DATA.map((entry, index) => (
          <React.Fragment key={index}>
            <ListItem alignItems="flex-start">
              <ListItemText
                primary={`Version ${entry.version} - ${entry.date}`}
                secondary={
                  <Box component="div" sx={{ mt: 1 }}>
                    {entry.changes.map((change, i) => (
                      <Box key={i} sx={{ display: 'flex', mb: 1 }}>
                        <Box sx={{ 
                          minWidth: 70, 
                          color: 
                            change.type === 'added' ? 'success.main' : 
                            change.type === 'fixed' ? 'warning.main' : 
                            'info.main' 
                        }}>
                          {change.type.charAt(0).toUpperCase() + change.type.slice(1)}:
                        </Box>
                        <Box sx={{ ml: 1 }}>{change.description}</Box>
                      </Box>
                    ))}
                  </Box>
                }
              />
            </ListItem>
            {index < CHANGELOG_DATA.length - 1 && <Divider component="li" />}
          </React.Fragment>
        ))}
      </List>
    </Box>
  );
};

export default VersionHistory; 