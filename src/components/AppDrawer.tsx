import React, { useState } from 'react';
import { Drawer } from '@mui/material';

const AppDrawer: React.FC = () => {
  const [open, setOpen] = useState(false);

  const handleClose = () => {
    setOpen(false);
  };

  return (
    <Drawer
      anchor="right"
      open={open}
      onClose={handleClose}
      sx={{
        '& .MuiDrawer-paper': {
          width: { xs: '100%', sm: 400 },
          maxWidth: '100%',
          overflowY: 'auto',
          padding: 2
        },
      }}
    >
      {/* Drawer content */}
    </Drawer>
  );
};

export default AppDrawer; 