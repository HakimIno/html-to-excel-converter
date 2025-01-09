const { execSync } = require('child_process');
const path = require('path');
const fs = require('fs');

function isDockerEnvironment() {
    try {
        return fs.existsSync('/.dockerenv') || fs.readFileSync('/proc/1/cgroup', 'utf8').includes('docker');
    } catch (error) {
        return false;
    }
}

function isProductionEnvironment() {
    return process.env.NODE_ENV === 'production';
}

function shouldSkipInstall() {
    // Skip if explicitly set to skip
    if (process.env.SKIP_PYTHON_INSTALL === 'true') {
        console.log('Skipping Python dependencies installation due to SKIP_PYTHON_INSTALL=true');
        return true;
    }
    
    // Allow force install in production if needed
    if (process.env.FORCE_PYTHON_INSTALL === 'true') {
        console.log('Forcing Python dependencies installation due to FORCE_PYTHON_INSTALL=true');
        return false;
    }
    
    // Skip in production by default unless in Docker
    const isProd = isProductionEnvironment();
    const isDocker = isDockerEnvironment();
    
    if (isProd && !isDocker) {
        console.log('Skipping Python dependencies installation in production environment');
        return true;
    }
    
    return false;
}

function checkPythonDependency(pythonCmd, dependency) {
    try {
        execSync(`${pythonCmd} -c "import ${dependency}"`, { stdio: 'ignore' });
        return true;
    } catch (error) {
        return false;
    }
}

function installPythonDependencies() {
    try {
        if (shouldSkipInstall()) {
            return;
        }

        const isDocker = isDockerEnvironment();
        const rootDir = path.join(__dirname, '..');
        const requirementsPath = path.join(rootDir, 'src', 'python', 'requirements.txt');
        
        // Check if python3 is available
        let pythonCmd = 'python3';
        try {
            execSync(`${pythonCmd} --version`);
        } catch (error) {
            try {
                pythonCmd = 'python';
                execSync(`${pythonCmd} --version`);
            } catch (error) {
                console.error('Python is not installed. Please install Python 3.7 or higher.');
                process.exit(1);
            }
        }

        if (isDocker) {
            // In Docker, install globally without virtual environment
            try {
                execSync(`${pythonCmd} -m pip install --upgrade pip`, { stdio: 'inherit' });
                execSync(`${pythonCmd} -m pip install -r "${requirementsPath}"`, { stdio: 'inherit' });
                console.log('Python dependencies installed successfully in Docker environment!');
                return;
            } catch (error) {
                console.error('Failed to install Python dependencies in Docker:', error.message);
                process.exit(1);
            }
        }

        // Local development environment
        console.log('Setting up Python virtual environment...');
        const venvPath = path.join(rootDir, 'venv');
        
        // Create .env file with venv path
        fs.writeFileSync(path.join(rootDir, '.env'), `VENV_PATH=${venvPath}`);

        // Check if virtual environment exists
        const hasVenv = fs.existsSync(venvPath);
        
        // Check if dependencies are already installed
        if (hasVenv) {
            const pipCmd = process.platform === 'win32' ? 
                path.join(venvPath, 'Scripts', 'python.exe') : 
                path.join(venvPath, 'bin', 'python');
                
            const missingDeps = [];
            ['bs4', 'xlsxwriter', 'pandas', 'cssutils'].forEach(dep => {
                if (!checkPythonDependency(pipCmd, dep)) {
                    missingDeps.push(dep);
                }
            });
            
            if (missingDeps.length === 0) {
                console.log('All Python dependencies are already installed.');
                return;
            }
        }
        
        // Create virtual environment if it doesn't exist
        if (!hasVenv) {
            try {
                execSync(`${pythonCmd} -m venv "${venvPath}"`, { stdio: 'inherit' });
            } catch (error) {
                console.error('Failed to create virtual environment. Trying to install dependencies globally...');
                try {
                    execSync(`${pythonCmd} -m pip install -r "${requirementsPath}"`, { stdio: 'inherit' });
                    console.log('Python dependencies installed globally successfully!');
                    return;
                } catch (globalError) {
                    console.error('Failed to install dependencies globally:', globalError.message);
                    process.exit(1);
                }
            }
        }
        
        // Install dependencies in virtual environment
        const pipCommand = process.platform === 'win32' ? 
            `"${path.join(venvPath, 'Scripts', 'pip.exe')}"` : 
            `"${path.join(venvPath, 'bin', 'pip')}"`;
            
        try {
            execSync(`${pipCommand} install --upgrade pip`, { stdio: 'inherit' });
            execSync(`${pipCommand} install -r "${requirementsPath}"`, { stdio: 'inherit' });
            console.log('Python dependencies installed successfully in virtual environment!');
        } catch (error) {
            console.error('Failed to install Python dependencies:', error.message);
            process.exit(1);
        }
        
    } catch (error) {
        console.error('Error during post-install setup:', error.message);
        process.exit(1);
    }
}

installPythonDependencies(); 