# PowerPoint Remote Control

A mobile-friendly web interface for remote PowerPoint control using your phone as a remote. Control presentations with simple touch gestures.

<img src="usage.png" alt="Usage Example" width="400" height="auto" style="max-width: 100%; height: auto;">

## Features

### 🎯 Core Functionality
- **Slide Navigation**: Tap left/right sides of screen to advance/previous slides
- **PowerPoint Integration**: Direct control of PowerPoint application
- **Fullscreen Mode**: Immersive presentation experience

### 📱 Mobile Interface
- **Responsive Design**: Works on any mobile device
- **Touch Controls**: Large, easy-to-tap navigation areas
- **Visual Feedback**: Clear visual indicators for all actions
- **Network Access**: QR code for easy connection

### 🛠️ Control Options
- **Focus PowerPoint**: Bring PowerPoint to front
- **Start Slideshow**: Begin presentation (F5)
- **Exit Slideshow**: End presentation (Escape)
- **Fullscreen Mode**: Toggle web app fullscreen

## Installation

### Prerequisites
- Node.js (v22 or higher)
- npm
- PowerPoint (Microsoft Office)

### Setup
1. **Clone the repository**
   ```bash
   git clone git@github.com:kokarn/pptx-control.git
   cd pptx-control
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Start the server**
   ```bash
   node server.mjs
   ```

4. **Access the interface**
   - The server will display local and network URLs
   - Scan the QR code with your phone for easy access
   - Or manually navigate to the displayed URL

## Usage

### Basic Navigation
1. **Start PowerPoint** and open your presentation
2. **Access the web interface** on your phone
3. **Tap "Focus PowerPoint"** to bring PowerPoint to front
4. **Tap "Start Slideshow"** to begin the presentation
5. **Navigate slides** by tapping left/right sides of the screen

### Advanced Features
- **Fullscreen mode**: Tap "Fullscreen" for immersive experience
- **Network access**: Works across your local network

## API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/` | GET | Main interface |
| `/click` | GET | Trigger mouse click |
| `/click-left` | GET | Press left arrow key |
| `/click-right` | GET | Press right arrow key |
| `/focus-powerpoint` | GET | Focus PowerPoint application |
| `/start-slideshow` | GET | Start PowerPoint slideshow |
| `/exit-slideshow` | GET | Exit PowerPoint slideshow |

## Platform Support

### ✅ Supported
- **macOS**: Full support with AppleScript integration
- **Windows**: Basic support (PowerPoint control via keyboard)
- **Mobile Browsers**: iOS Safari, Chrome, Firefox
- **Desktop Browsers**: Chrome, Firefox, Safari, Edge

### 🔧 Requirements
- **PowerPoint**: Microsoft PowerPoint must be installed
- **Network**: Both devices must be on same network

## Development

### Project Structure
```
pptx-control/
├── server.mjs          # Express server with endpoints
├── public/
│   └── index.html      # Main mobile interface
├── package.json        # Dependencies and scripts
└── README.md          # This file
```

### Key Technologies
- **Backend**: Node.js, Express.js
- **Frontend**: HTML5, CSS3, JavaScript
- **Automation**: RobotJS for system control
- **QR Codes**: qrcode-terminal for easy access

### Contributing
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is open source and available under the MIT License.

## Support

For issues, questions, or contributions:
- Create an issue on GitHub

---

**Happy presenting! 🎤📊**
