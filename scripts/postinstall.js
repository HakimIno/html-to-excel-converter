const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');

function setupPythonEnvironment() {
    try {
        const isWindows = os.platform() === 'win32';
        const pythonCmd = isWindows ? 'python' : 'python3';
        const srcPath = path.join(__dirname, '..', 'src');
        const venvPath = path.join(srcPath, 'venv');
        const requirementsPath = path.join(srcPath, 'python', 'requirements.txt');
        const markerPath = path.join(venvPath, '.installation_complete');

        // Check if already installed
        if (fs.existsSync(markerPath)) {
            console.log('Python environment already set up.');
            return;
        }

        console.log('Setting up Python environment...');

        // Create virtual environment
        execSync(`${pythonCmd} -m venv "${venvPath}"`, { stdio: 'inherit' });

        // Get pip path
        const pipCmd = isWindows ? 
            path.join(venvPath, 'Scripts', 'pip.exe') :
            path.join(venvPath, 'bin', 'pip');

        // Install requirements
        console.log('Installing Python dependencies...');
        execSync(`"${pipCmd}" install -r "${requirementsPath}"`, { stdio: 'inherit' });

        // Create marker file
        fs.writeFileSync(markerPath, new Date().toISOString());
        console.log('Python environment setup complete.');

    } catch (error) {
        console.error('Failed to setup Python environment:', error.message);
        process.exit(1);
    }
}

setupPythonEnvironment(); 