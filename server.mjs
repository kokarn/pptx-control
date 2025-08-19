import os from 'os';
import { exec } from 'child_process';

import express from 'express';
import robot from 'robotjs';
import qrcode from 'qrcode-terminal';

const app = express();
const port = 3000;

// Function to get local network IP address
function getLocalIP() {
  const interfaces = os.networkInterfaces();
  for (const name of Object.keys(interfaces)) {
    for (const iface of interfaces[name]) {
      // Skip internal and non-IPv4 addresses
      if (iface.family === 'IPv4' && !iface.internal) {
        return iface.address;
      }
    }
  }
  return 'localhost';
}

// Function to focus on PowerPoint application
function focusPowerPoint() {
  return new Promise((resolve, reject) => {
    // On macOS, use osascript to focus on PowerPoint
    if (process.platform === 'darwin') {
      exec('osascript -e \'tell application "Microsoft PowerPoint" to activate\'', (error, stdout, stderr) => {
        if (error) {
          console.error('Error focusing PowerPoint:', error);
          reject(error);
        } else {
          console.log('PowerPoint focused successfully');
          resolve();
        }
      });
    } else {
      // On Windows, you might need a different approach
      console.log('PowerPoint focus not implemented for this platform');
      resolve();
    }
  });
}

// Function to start PowerPoint slideshow
function startSlideshow() {
  return new Promise((resolve, reject) => {
    try {
      // Use F5 key to start slideshow (standard PowerPoint shortcut)
      robot.keyTap('f5');
      console.log('Slideshow started successfully');
      resolve();
    } catch (error) {
      console.error('Error starting slideshow:', error);
      reject(error);
    }
  });
}

// Serve static files from public directory
app.use(express.static('public'));

app.get('/click', (req, res) => {
  robot.mouseClick();
  // Return 'Clicked' with a timestamp
  console.log(`Clicked at ${new Date().toISOString()}`);
  res.send(`'Clicked' ${new Date().toISOString()}`);
});

app.get('/click-left', (req, res) => {
  robot.keyTap('left');
  // Return 'Left arrow pressed' with a timestamp
  console.log(`Left arrow pressed at ${new Date().toISOString()}`);
  res.send(`'Left arrow pressed' ${new Date().toISOString()}`);
});

app.get('/click-right', (req, res) => {
  robot.keyTap('right');
  // Return 'Right arrow pressed' with a timestamp
  console.log(`Right arrow pressed at ${new Date().toISOString()}`);
  res.send(`'Right arrow pressed' ${new Date().toISOString()}`);
});

app.get('/focus-powerpoint', async (req, res) => {
  try {
    await focusPowerPoint();
    console.log(`PowerPoint focused at ${new Date().toISOString()}`);
    res.send(`'PowerPoint focused' ${new Date().toISOString()}`);
  } catch (error) {
    console.error('Failed to focus PowerPoint:', error);
    res.status(500).send(`'Failed to focus PowerPoint' ${new Date().toISOString()}`);
  }
});

app.get('/start-slideshow', async (req, res) => {
  try {
    await focusPowerPoint();
    // Wait a moment for PowerPoint to focus
    await new Promise(resolve => setTimeout(resolve, 500));
    await startSlideshow();
    console.log(`Slideshow started at ${new Date().toISOString()}`);
    res.send(`'Slideshow started' ${new Date().toISOString()}`);
  } catch (error) {
    console.error('Failed to start slideshow:', error);
    res.status(500).send(`'Failed to start slideshow' ${new Date().toISOString()}`);
  }
});

app.get('/exit-slideshow', async (req, res) => {
  try {
    await focusPowerPoint();
    // Wait a moment for PowerPoint to focus
    await new Promise(resolve => setTimeout(resolve, 500));
    // Press Escape key to exit slideshow
    robot.keyTap('escape');
    console.log(`Slideshow exited at ${new Date().toISOString()}`);
    res.send(`'Slideshow exited' ${new Date().toISOString()}`);
  } catch (error) {
    console.error('Failed to exit slideshow:', error);
    res.status(500).send(`'Failed to exit slideshow' ${new Date().toISOString()}`);
  }
});

// Start the server
app.listen(port, () => {
  const localIP = getLocalIP();
  console.log(`Server is running at:`);
  console.log(`  Local: http://localhost:${port}`);
  console.log(`  Network: http://${localIP}:${port}`);
  qrcode.generate(`http://${localIP}:${port}`);
});
