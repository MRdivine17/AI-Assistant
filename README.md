<<<<<<< HEAD
# 🎯 Jarvis Desktop Voice Assistant

<div align="center">
  <img src="https://giffiles.alphacoders.com/212/212508.gif" alt="Jarvis Assistant" width="400">
  
  [![Python](https://img.shields.io/badge/Python-3.6+-blue.svg)](https://python.org)
  [![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
  [![Windows](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)
</div>

### BEFORE YOU RUN THESE CHANGE THE API KEY OF THE OPENAI, GEMINI AND GROK IN THE OPEN ROUNTER 

TO RUN THIS USE THIS IN THE POWERSHELL OR TERMINAL 
```
cd C:\Users\USER\Desktop\personal_assistant; taskkill /im python.exe /f; .\.venv\Scripts\python Jarvis-Desktop-Voice-Assistant\Jarvis\jarvis.py
```
USE THIS COMMAND 

## 📋 Overview

**Jarvis Desktop Voice Assistant** is an intelligent voice-controlled desktop assistant built with Python that can perform various tasks through voice commands. This project brings the concept of a personal AI assistant to your desktop, making daily computer tasks more convenient and automated.

### 🌟 Key Features

- **🎤 Voice Recognition**: Listen and respond to voice commands
- **🗣️ Text-to-Speech**: Natural voice responses with multiple voice options
- **🌐 Web Integration**: Search Wikipedia, Google, and open websites
- **🎵 Media Control**: Play music locally or online, control playback
- **📁 File Management**: Create, delete, rename files and folders
- **💻 System Control**: Launch applications, take screenshots, system operations
- **🤖 AI Integration**: Support for multiple AI providers (OpenAI, Google Gemini, Anthropic)
- **⌨️ Multiple Input Methods**: Voice commands, keyboard input, or typed commands
- **🎨 Modern GUI**: Dark-themed interface with real-time conversation logging
- **🔄 Wake Word Detection**: Always-on listening for "Jarvis" wake word
- **🚀 Auto-Startup**: Optional Windows startup integration

## 🛠️ Technologies Used

- **Python 3.6+** - Core programming language
- **Speech Recognition** - Voice input processing
- **pyttsx3** - Text-to-speech conversion
- **Wikipedia API** - Information retrieval
- **PyAutoGUI** - System automation and screenshots
- **Tkinter** - GUI interface
- **VLC Media Player** - Audio playback
- **yt-dlp** - YouTube audio streaming
- **Multiple AI APIs** - OpenAI, Google Gemini, Anthropic integration

## 📦 Installation & Setup

### Prerequisites
- Python 3.6 or higher
- Windows operating system
- Microphone and speakers/headphones
- Internet connection (for AI features and web searches)

### Step-by-Step Installation

1. **Clone the Repository**
   ```bash
   git clone https://github.com/MRdivine17/Jarvis-PC-assistant.git
   cd Jarvis-Desktop-Voice-Assistant
   ```

2. **Create Virtual Environment**
   ```bash
   python -m venv jarvis_env
   jarvis_env\Scripts\activate  # Windows
   ```

3. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Install PyAudio** (if needed)
   - Download the appropriate wheel file from [here](https://www.lfd.uci.edu/~gohlke/pythonlibs/#pyaudio)
   - Install using: `pip install PyAudio-0.2.11-cp39-cp39-win_amd64.whl`

5. **Configure AI APIs (Optional)**
   Create a `.env` file in the project root:
   ```env
   # Choose your preferred AI provider
   LLM_PROVIDER=openrouter  # or openai, gemini, anthropic
   
   # OpenRouter (recommended - free tier available)
   OPENROUTER_API_KEY=your_api_key_here
   OPENROUTER_MODEL=openai/gpt-4o
   
   # Or use other providers
   OPENAI_API_KEY=your_openai_key
   GOOGLE_API_KEY=your_google_key
   ANTHROPIC_API_KEY=your_anthropic_key
   ```

6. **Run the Assistant**
   ```bash
   python Jarvis/jarvis.py
   ```

## 🎮 Usage

### Voice Commands

#### Basic Commands
- **"Jarvis"** - Wake word to activate
- **"What time is it?"** - Get current time
- **"What's the date?"** - Get current date
- **"Tell me a joke"** - Hear a random joke

#### Web & Search
- **"Search [query] on Google"** - Google search
- **"Search [topic] on Wikipedia"** - Wikipedia search
- **"Open YouTube"** - Launch YouTube
- **"Open [website]"** - Open any website

#### Media Control
- **"Play music"** - Play local music
- **"Play [song name] online"** - Stream music from YouTube
- **"Pause music"** - Pause current playback
- **"Stop music"** - Stop playback
- **"Next song"** - Skip to next track

#### File Management
- **"Create folder [name]"** - Create new folder
- **"Delete folder [name]"** - Delete folder
- **"Open folder [name]"** - Navigate to folder
- **"List files"** - Show current directory contents

#### System Control
- **"Open [application]"** - Launch applications
- **"Take screenshot"** - Capture screen
- **"Shutdown"** - Shutdown computer
- **"Restart"** - Restart computer

#### AI Features
- **"What is [topic]?"** - Get AI-powered explanations
- **"How do I [task]?"** - Get step-by-step instructions
- **"Explain [concept]"** - Detailed explanations

### GUI Controls

- **Press 'N'** - Push-to-talk
- **Press 'S'** - Stop current speech
- **Checkbox** - Toggle wake word detection
- **Text Input** - Type commands directly
- **Audio Test** - Test voice output

## 🔧 Configuration

### Voice Settings
- **Change Voice**: "Change voice to [number]" or "Use voice [number]"
- **List Voices**: "List voices" to see available options
- **Voice Types**: "Use female voice" or "Use male voice"

### AI Provider Setup
The assistant supports multiple AI providers:

1. **OpenRouter** (Recommended)
   - Free tier available
   - Access to multiple models
   - Set `LLM_PROVIDER=openrouter`

2. **OpenAI**
   - Set `LLM_PROVIDER=openai`
   - Requires OpenAI API key

3. **Google Gemini**
   - Set `LLM_PROVIDER=gemini`
   - Requires Google API key

4. **Anthropic Claude**
   - Set `LLM_PROVIDER=anthropic`
   - Requires Anthropic API key

## 📁 Project Structure

```
Jarvis-Desktop-Voice-Assistant/
├── Jarvis/
│   ├── jarvis.py          # Main application file
│   ├── data.txt           # Assistant data storage
│   └── README.md          # Detailed documentation
├── Images/                # Screenshots and assets
├── Presentation/          # Project presentation files
├── requirements.txt       # Python dependencies
├── LICENSE               # MIT License
└── README.md            # This file
```

## 🚀 Advanced Features

### Wake Word Detection
- Always-on listening for "Jarvis"
- Configurable wake word sensitivity
- Background processing with minimal resource usage

### Multi-Provider AI Support
- Automatic fallback between AI providers
- Model selection and configuration
- Context-aware responses

### File System Integration
- Smart file and folder detection
- Cross-directory search capabilities
- Windows registry integration for app launching

### Media Streaming
- YouTube audio streaming
- Local music playback
- VLC integration for high-quality audio

## 🛡️ Security & Privacy

- **Local Processing**: Voice recognition and basic commands run locally
- **API Keys**: Securely stored in environment variables
- **No Data Collection**: No personal data is stored or transmitted
- **Open Source**: Full source code available for review

## 🐛 Troubleshooting

### Common Issues

1. **Microphone Not Working**
   - Check microphone permissions
   - Ensure microphone is set as default device
   - Test with Windows Speech Recognition

2. **PyAudio Installation Failed**
   - Install Microsoft Visual C++ Build Tools
   - Use pre-compiled wheel files
   - Try alternative installation methods

3. **Voice Recognition Issues**
   - Speak clearly and at normal volume
   - Reduce background noise
   - Check internet connection for Google Speech API

4. **AI Features Not Working**
   - Verify API keys in `.env` file
   - Check internet connection
   - Ensure sufficient API credits

### Performance Optimization

- **Reduce Wake Word Sensitivity**: Disable if not needed
- **Limit AI Usage**: Use local commands when possible
- **Close Unused Applications**: Free up system resources



1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Development Setup
```bash
# Install development dependencies
pip install -r requirements.txt
pip install pytest black flake8

# Run tests
pytest

# Format code
black Jarvis/jarvis.py
```

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👨‍💻 Author


- GitHub: [@MRdivine17](https://github.com/MRdivine17)

## 🙏 Acknowledgments

- Google Speech Recognition API
- Wikipedia for information access
- VLC Media Player for audio support
- All AI providers (OpenAI, Google, Anthropic)
- Python community for excellent libraries



## ⭐ Show Your Support

If this project helped you, please give it a ⭐️ and share it with others!

---

<div align="center">
  <strong>Made with ❤️ by Kishan Kumar Rai</strong>
</div>

