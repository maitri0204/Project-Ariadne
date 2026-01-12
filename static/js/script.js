let currentFilename = null;

// Theme Toggle Functionality
document.addEventListener('DOMContentLoaded', function() {
    const themeToggle = document.getElementById('themeToggle');
    const themeIcon = document.getElementById('themeIcon');
    const themeText = document.getElementById('themeText');
    const body = document.body;

    // Check for saved theme preference or default to light mode
    const currentTheme = localStorage.getItem('theme') || 'light';
    
    // Apply saved theme
    if (currentTheme === 'dark') {
        body.classList.add('dark-theme');
        themeIcon.classList.remove('fa-moon');
        themeIcon.classList.add('fa-sun');
        themeText.textContent = 'Light Mode';
    }

    // Theme toggle click event
    themeToggle.addEventListener('click', function() {
        body.classList.toggle('dark-theme');
        
        if (body.classList.contains('dark-theme')) {
            themeIcon.classList.remove('fa-moon');
            themeIcon.classList.add('fa-sun');
            themeText.textContent = 'Light Mode';
            localStorage.setItem('theme', 'dark');
        } else {
            themeIcon.classList.remove('fa-sun');
            themeIcon.classList.add('fa-moon');
            themeText.textContent = 'Dark Mode';
            localStorage.setItem('theme', 'light');
        }
    });
});

function generateReport(reportType) {
    // Collect form data
    const formData = collectFormData();
    
    // Validate that at least some selections are made
    if (!validateFormData(formData)) {
        return;
    }

    // Show loading spinner
    document.getElementById('loadingSpinner').style.display = 'block';
    document.getElementById('reportSection').style.display = 'none';

    // Scroll to loading spinner
    document.getElementById('loadingSpinner').scrollIntoView({ behavior: 'smooth' });

    // Prepare request data
    const requestData = {
        report_type: reportType,
        inputs: formData
    };

    // Make API call
    fetch('/generate-report', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(requestData)
    })
    .then(response => response.json())
    .then(data => {
        // Hide loading spinner
        document.getElementById('loadingSpinner').style.display = 'none';

        if (data.error) {
            alert('Error: ' + data.error);
            return;
        }

        // Display report
        displayReport(data.content, data.filename);
    })
    .catch(error => {
        document.getElementById('loadingSpinner').style.display = 'none';
        alert('Error generating report: ' + error.message);
        console.error('Error:', error);
    });
}

// Toggle percentage input visibility when checkbox is selected
// Toggle percentage input visibility when checkbox is selected
function togglePercentageInput(checkbox, category) {
    const value = checkbox.value;
    const percentageInput = document.getElementById(`percentage-${category}-${value}`);
    
    if (percentageInput) {
        if (checkbox.checked) {
            percentageInput.style.display = 'block';
            percentageInput.required = true;
            
            // Add real-time validation to prevent values > 100
            percentageInput.addEventListener('input', function() {
                // Remove any non-numeric characters except decimal point
                this.value = this.value.replace(/[^0-9.]/g, '');
                
                // Prevent multiple decimal points
                const parts = this.value.split('.');
                if (parts.length > 2) {
                    this.value = parts[0] + '.' + parts.slice(1).join('');
                }
                
                // Limit to 100
                if (parseFloat(this.value) > 100) {
                    this.value = '100';
                }
                
                // Limit decimal places to 1
                if (parts.length === 2 && parts[1].length > 1) {
                    this.value = parseFloat(this.value).toFixed(1);
                }
            });
            
        } else {
            percentageInput.style.display = 'none';
            percentageInput.required = false;
            percentageInput.value = ''; // Clear value when unchecked
        }
    }
}

function collectFormData() {
    const formData = {};
    // Collect standard + board
    formData.sname = document.getElementById("sname").value.trim();
    formData.standard = document.getElementById("standard").value.trim();
    formData.board = document.getElementById("board").value;
    
    // Collect highest skills with percentages
    const highest_Skills = [];
    const skillPercentages = {};
    document.querySelectorAll('input[name="highest_skills"]:checked').forEach(checkbox => {
        const skill = checkbox.value;
        highest_Skills.push(skill);
        const percentageInput = document.getElementById(`percentage-skill-${skill}`);
        if (percentageInput && percentageInput.value) {
            skillPercentages[skill] = parseFloat(percentageInput.value);
        }
    });
    formData.highest_skills = highest_Skills;
    formData.skillpercentages = skillPercentages;
    
    // Collect thinking pattern (radio - no percentage needed)
    const thinking_Pattern = document.querySelector('input[name="thinking_pattern"]:checked');
    formData.thinking_pattern = thinking_Pattern ? thinking_Pattern.value : '';
    
    // Collect achievement style with percentages
    const achievement_Style = [];
    const achievementPercentages = {};
    document.querySelectorAll('input[name="achievement_style"]:checked').forEach(checkbox => {
        const style = checkbox.value;
        achievement_Style.push(style);
        const percentageInput = document.getElementById(`percentage-achievement-${style}`);
        if (percentageInput && percentageInput.value) {
            achievementPercentages[style] = parseFloat(percentageInput.value);
        }
    });
    formData.achievement_style = achievement_Style;
    formData.achievementpercentages = achievementPercentages;
    
    // Collect learning communication style with percentages
    const learning_Communication_Style = [];
    const learningPercentages = {};
    document.querySelectorAll('input[name="learning_communication_style"]:checked').forEach(checkbox => {
        const style = checkbox.value;
        learning_Communication_Style.push(style);
        const percentageInput = document.getElementById(`percentage-learning-${style}`);
        if (percentageInput && percentageInput.value) {
            learningPercentages[style] = parseFloat(percentageInput.value);
        }
    });
    formData.learning_communication_style = learning_Communication_Style;
    formData.learningpercentages = learningPercentages;
    
    // Collect quotients with percentages
    const quotients = [];
    const quotientPercentages = {};
    document.querySelectorAll('input[name="quotients"]:checked').forEach(checkbox => {
        const quotient = checkbox.value;
        quotients.push(quotient);
        const percentageInput = document.getElementById(`percentage-quotient-${quotient}`);
        if (percentageInput && percentageInput.value) {
            quotientPercentages[quotient] = parseFloat(percentageInput.value);
        }
    });
    formData.quotients = quotients;
    formData.quotientpercentages = quotientPercentages;
    
    // Collect personality type (radio - no percentage needed)
    const personality_Type = document.querySelector('input[name="personality_type"]:checked');
    formData.personality_type = personality_Type ? personality_Type.value : '';
    
    // Collect career roles
    formData.career_roles = document.getElementById('careerRoles').value.trim();
    
    return formData;
}

function validateFormData(formData) {

    if (!formData.sname) {
        alert("Please enter the student's name.");
        return false;
    }

    if (!formData.standard) {
        alert("Please enter the student's standard / year.");
        return false;
    }
      
    if (!formData.board) {
        alert("Please select the board.");
        return false;
    }
      
    // Check if at least one skill is selected
    if (formData.highest_skills.length === 0) {
        alert('Please select at least one skill');
        return false;
    }
    
    // Validate that selected skills have percentages and are within range
    for (let skill of formData.highest_skills) {
        const percentage = formData.skillpercentages[skill];
        if (!percentage || percentage === 0) {
            alert(`⚠️ Please enter percentage for skill: ${skill}`);
            return false;
        }
        if (percentage > 100) {
            alert(`⚠️ Percentage for ${skill} cannot exceed 100%`);
            return false;
        }
    }
    
    // Validate achievement style percentages if selected
    for (let style of formData.achievement_style) {
        const percentage = formData.achievementpercentages[style];
        if (!percentage || percentage === 0) {
            alert(`⚠️ Please enter percentage for achievement style: ${style}`);
            return false;
        }
        if (percentage > 100) {
            alert(`⚠️ Percentage for ${style} cannot exceed 100%`);
            return false;
        }
    }

    // Validate learning style percentages if selected
    for (let style of formData.learning_communication_style) {
        const percentage = formData.learningpercentages[style];
        if (!percentage || percentage === 0) {
            alert(`⚠️ Please enter percentage for learning style: ${style}`);
            return false;
        }
        if (percentage > 100) {
            alert(`⚠️ Percentage for ${style} cannot exceed 100%`);
            return false;
        }
    }

    // Validate quotient percentages if selected
    for (let quotient of formData.quotients) {
        const percentage = formData.quotientpercentages[quotient];
        if (!percentage || percentage === 0) {
            alert(`⚠️ Please enter percentage for quotient: ${quotient}`);
            return false;
        }
        if (percentage > 100) {
            alert(`⚠️ Percentage for ${quotient} cannot exceed 100%`);
            return false;
        }
    }
    
    // Check if thinking pattern is selected
    if (!formData.thinking_pattern) {
        alert('Please select a thinking pattern');
        return false;
    }
    
    // Check if personality type is selected
    if (!formData.personality_type) {
        alert('Please select a personality type');
        return false;
    }
    
    return true;
}

function displayReport(content, filename) {
    currentFilename = filename;

    // Format content with proper HTML
    const formattedContent = formatReportContent(content);

    // Display report section
    const reportContent = document.getElementById('reportContent');
    reportContent.innerHTML = formattedContent;

    const reportSection = document.getElementById('reportSection');
    reportSection.style.display = 'block';

    // Scroll to report
    reportSection.scrollIntoView({ behavior: 'smooth' });

    // Set up download button
    const downloadBtn = document.getElementById('downloadBtn');
    downloadBtn.onclick = function() {
        downloadReport(filename);
    };
}

function formatReportContent(content) {
    // Convert markdown-style content to HTML
    let formatted = content;

    // Convert headings
    formatted = formatted.replace(/###\s+(.+)/g, '<h4>$1</h4>');
    formatted = formatted.replace(/##\s+(.+)/g, '<h3>$1</h3>');
    formatted = formatted.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');

    // Convert bullet points
    formatted = formatted.replace(/^\-\s+(.+)$/gm, '<li>$1</li>');
    formatted = formatted.replace(/^•\s+(.+)$/gm, '<li>$1</li>');

    // Wrap consecutive <li> in <ul>
    formatted = formatted.replace(/(<li>.*?<\/li>\s*)+/gs, match => {
        return '<ul>' + match + '</ul>';
    });

    // Convert line breaks to paragraphs
    formatted = formatted.split('\n\n').map(para => {
        para = para.trim();
        if (!para) return '';
        if (!para.startsWith('<')) {
            return '<p>' + para + '</p>';
        }
        return para;
    }).join('\n');

    return formatted;
}

function downloadReport(filename) {
    window.location.href = '/download/' + filename;
}

// Add smooth scrolling to all links
document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function (e) {
        e.preventDefault();
        const target = document.querySelector(this.getAttribute('href'));
        if (target) {
            target.scrollIntoView({
                behavior: 'smooth'
            });
        }
    });
});

// Add animation to form elements on focus
document.querySelectorAll('input, textarea').forEach(element => {
    element.addEventListener('focus', function() {
        this.style.transform = 'scale(1.02)';
    });

    element.addEventListener('blur', function() {
        this.style.transform = 'scale(1)';
    });
});
