import { Document, Paragraph, TextRun, Packer, AlignmentType } from 'https://cdn.jsdelivr.net/npm/docx@8.5.0/+esm';

class SyllabusBuilder {
    constructor() {
        this.form = document.getElementById('syllabusForm');
        this.sections = [
            { id: 'basic-info', requiredFields: ['courseTitle', 'courseNumber', 'term', 'credits', 'instructorName', 'officeHours', 'classSchedule'] },
            { id: 'description', requiredFields: ['courseDescription'] },
            { id: 'learning-outcomes', requiredFields: ['learningOutcomes'] },
            { id: 'communication', optionalFields: ['communicationStrategy'] },
            { id: 'technical', optionalFields: ['technicalRequirements'] },
            { id: 'requirements', requiredFields: ['assignmentTypes', 'gradingPercentages', 'dueDatesPolicy'] },
            { id: 'outline', customValidator: () => this.validateCourseOutline() },
            { id: 'integrity', requiredFields: ['academicIntegrity'] },
            { id: 'code-of-conduct', requiredFields: ['codeOfConduct'] },
            { id: 'integrity-credit', requiredFields: ['integrityOfCredit'] },
            { id: 'conduct', requiredFields: ['generalConduct'] },
            { id: 'grading', requiredFields: ['gradingGuidelines'] },
            { id: 'accessibility', requiredFields: ['studentAccessibility'] },
            { id: 'wellness', optionalFields: ['studentWellness'] },
            { id: 'pronouns', requiredFields: ['namePronouns'] }, // Pre-filled, always complete
            { id: 'religious', optionalFields: ['religiousObservances'] },
            { id: 'devices', optionalFields: ['electronicDevices'] },
            { id: 'ai', optionalFields: ['aiGuidance'] }
        ];

        this.initializeEventListeners();
        this.initializeCourseOutlineBuilder();
        this.initializeTOC();
        this.updateTOCStatus();
    }

    initializeEventListeners() {
        document.getElementById('loadSampleBtn').addEventListener('click', () => this.loadSampleData());
        document.getElementById('loadBtn').addEventListener('click', () => this.loadSyllabus());
        document.getElementById('saveBtn').addEventListener('click', () => this.saveSyllabus());
        document.getElementById('exportWordBtn').addEventListener('click', () => this.exportToWord());
        document.getElementById('exportPdfBtn').addEventListener('click', () => this.exportToPDF());

        // Handle file input for loading
        document.getElementById('loadFile').addEventListener('change', (e) => this.handleFileLoad(e));

        // Update TOC status when form fields change
        this.form.addEventListener('input', () => this.updateTOCStatus());
        this.form.addEventListener('change', () => this.updateTOCStatus());
    }

    initializeCourseOutlineBuilder() {
        // Add module button handler
        const addModuleBtn = document.getElementById('addModuleBtn');
        if (addModuleBtn) {
            addModuleBtn.onclick = () => this.addModule();
        }

        // Set up event delegation for all buttons (one-time setup)
        this.setupRemoveHandlers();

        // Initialize with one module containing one class day
        this.addModule();
    }

    addModule() {
        const container = document.getElementById('moduleContainer');
        const moduleItem = document.createElement('div');
        const moduleNumber = container.children.length + 1;
        const moduleId = 'module-' + Date.now();
        moduleItem.className = 'module-item';
        moduleItem.dataset.moduleId = moduleId;
        moduleItem.innerHTML = `
            <div class="module-header">
                <input type="text" class="module-title" placeholder="Module ${moduleNumber}: Introduction to Topic">
                <button type="button" class="remove-module-btn">Remove Module</button>
            </div>
            <textarea class="module-description" rows="2" placeholder="Brief module description or learning objectives (optional)"></textarea>
            <div class="module-class-days" data-module-id="${moduleId}">
                <!-- Class days for this module -->
            </div>
            <button type="button" class="add-class-day-btn" data-module-id="${moduleId}">Add Class Day</button>
        `;
        container.appendChild(moduleItem);

        // Add one initial class day to the new module
        this.addClassDay(moduleId);

        this.updateTOCStatus();
    }

    addClassDay(moduleId) {
        const container = document.querySelector(`.module-class-days[data-module-id="${moduleId}"]`);
        if (!container) return;

        const classDayItem = document.createElement('div');
        const dayNumber = container.children.length + 1;
        classDayItem.className = 'class-day-item';
        classDayItem.innerHTML = `
            <div class="class-day-header">
                <input type="text" class="class-day-title" placeholder="Class Day ${dayNumber}: Topic Name" required>
                <button type="button" class="remove-class-btn">Remove</button>
            </div>
            <textarea class="class-day-content" rows="6" placeholder="Readings: List your readings here.&#10;&#10;Activity: Describe class activities.&#10;&#10;Assignment: Assignment description and due date." required></textarea>
        `;
        container.appendChild(classDayItem);
        this.updateTOCStatus();
    }

    setupRemoveHandlers() {
        // Use event delegation for all buttons (one-time setup)
        const moduleContainer = document.getElementById('moduleContainer');
        if (moduleContainer) {
            moduleContainer.addEventListener('click', (e) => {
                // Handle "Add Class Day" button clicks
                if (e.target.classList.contains('add-class-day-btn')) {
                    const moduleId = e.target.dataset.moduleId;
                    this.addClassDay(moduleId);
                }

                // Handle "Remove Module" button clicks
                if (e.target.classList.contains('remove-module-btn')) {
                    const moduleItems = moduleContainer.querySelectorAll('.module-item');
                    if (moduleItems.length > 1) {
                        e.target.closest('.module-item').remove();
                        this.updateTOCStatus();
                    } else {
                        alert('You must have at least one module.');
                    }
                }

                // Handle "Remove Class Day" button clicks
                if (e.target.classList.contains('remove-class-btn')) {
                    const moduleItem = e.target.closest('.module-item');
                    const classDaysContainer = moduleItem.querySelector('.module-class-days');
                    const classDayItems = classDaysContainer.querySelectorAll('.class-day-item');

                    if (classDayItems.length > 1) {
                        e.target.closest('.class-day-item').remove();
                        this.updateTOCStatus();
                    } else {
                        alert('Each module must have at least one class day.');
                    }
                }
            });
        }
    }

    validateCourseOutline() {
        // Check if at least one class day has title and content filled
        const classDays = document.querySelectorAll('.class-day-item');
        for (let day of classDays) {
            const title = day.querySelector('.class-day-title').value.trim();
            const content = day.querySelector('.class-day-content').value.trim();
            if (title && content) {
                return true;
            }
        }
        return false;
    }

    collectCourseOutlineData() {
        // Collect modules with their nested class days
        const modules = [];
        document.querySelectorAll('.module-item').forEach(item => {
            const title = item.querySelector('.module-title').value.trim();
            const description = item.querySelector('.module-description').value.trim();

            // Collect class days for this module
            const classDays = [];
            const classDaysContainer = item.querySelector('.module-class-days');
            if (classDaysContainer) {
                classDaysContainer.querySelectorAll('.class-day-item').forEach(dayItem => {
                    const dayTitle = dayItem.querySelector('.class-day-title').value.trim();
                    const dayContent = dayItem.querySelector('.class-day-content').value.trim();
                    if (dayTitle && dayContent) {
                        classDays.push({ title: dayTitle, content: dayContent });
                    }
                });
            }

            // Only add module if it has content
            if (title || classDays.length > 0) {
                modules.push({
                    title,
                    description,
                    classDays
                });
            }
        });

        return { modules };
    }

    generateCourseOutlineHTML(outlineData) {
        let html = '';

        // Iterate through modules
        if (outlineData.modules && outlineData.modules.length > 0) {
            outlineData.modules.forEach(module => {
                // Module header
                if (module.title) {
                    html += `<div class="course-outline-module">
                        <h4>${module.title}</h4>`;
                    if (module.description) {
                        html += `<p>${module.description}</p>`;
                    }
                    html += `</div>`;
                }

                // Class days table for this module
                if (module.classDays && module.classDays.length > 0) {
                    html += `<table class="course-outline-table">
                        <thead>
                            <tr>
                                <th>Class Day & Title</th>
                                <th>Readings, Assignments, and Activities</th>
                            </tr>
                        </thead>
                        <tbody>`;

                    module.classDays.forEach(day => {
                        const formattedContent = day.content.replace(/\n/g, '<br>');
                        html += `<tr>
                            <td class="class-day-title-cell"><strong>${day.title}</strong></td>
                            <td class="class-day-content-cell">${formattedContent}</td>
                        </tr>`;
                    });

                    html += `</tbody></table>`;
                }
            });
        }

        return html;
    }

    formatCourseOutlineForExport(outlineData) {
        let text = '';

        // Iterate through modules with their class days
        if (outlineData.modules && outlineData.modules.length > 0) {
            outlineData.modules.forEach((module, moduleIndex) => {
                if (moduleIndex > 0) text += '\n\n';

                // Module header
                if (module.title) {
                    text += `${module.title}\n`;
                    if (module.description) {
                        text += `${module.description}\n`;
                    }
                    text += '\n';
                }

                // Class days for this module
                if (module.classDays && module.classDays.length > 0) {
                    module.classDays.forEach((day, dayIndex) => {
                        if (dayIndex > 0) text += '\n';
                        text += `${day.title}\n${day.content}\n`;
                    });
                }
            });
        }

        return text;
    }

    initializeTOC() {
        // Add click handlers for TOC links
        const tocLinks = document.querySelectorAll('.sidebar-toc a');
        tocLinks.forEach(link => {
            link.addEventListener('click', (e) => {
                // Remove active class from all links
                tocLinks.forEach(l => l.classList.remove('active'));
                // Add active class to clicked link
                e.currentTarget.classList.add('active');
            });
        });

        // Highlight active section on scroll
        const observerOptions = {
            root: null,
            rootMargin: '-100px 0px -70% 0px',
            threshold: 0
        };

        const observer = new IntersectionObserver((entries) => {
            entries.forEach(entry => {
                if (entry.isIntersecting) {
                    const id = entry.target.getAttribute('id');
                    tocLinks.forEach(link => {
                        if (link.getAttribute('data-section') === id) {
                            tocLinks.forEach(l => l.classList.remove('active'));
                            link.classList.add('active');
                        }
                    });
                }
            });
        }, observerOptions);

        // Observe all sections
        this.sections.forEach(section => {
            const element = document.getElementById(section.id);
            if (element) {
                observer.observe(element);
            }
        });
    }

    updateTOCStatus() {
        this.sections.forEach(section => {
            const link = document.querySelector(`a[data-section="${section.id}"]`);
            if (!link) return;

            // Remove existing status classes
            link.classList.remove('complete', 'incomplete', 'optional');

            // Helper function to check if a field is filled
            const isFieldFilled = (fieldId) => {
                const field = document.getElementById(fieldId);
                if (!field) return true; // Field doesn't exist, consider it filled

                if (field.type === 'checkbox') {
                    return field.checked;
                } else if (field.tagName === 'SELECT') {
                    return field.value !== '';
                } else if (field.hasAttribute('readonly')) {
                    // Readonly fields are pre-filled, always consider them complete
                    return field.value.trim() !== '';
                } else {
                    return field.value.trim() !== '';
                }
            };

            // Check if section has a custom validator
            if (section.customValidator) {
                const isValid = section.customValidator();
                if (isValid) {
                    link.classList.add('complete');
                } else {
                    link.classList.add('incomplete');
                }
            }
            // Check if section has required fields
            else if (section.requiredFields && section.requiredFields.length > 0) {
                // Required section
                const allFilled = section.requiredFields.every(isFieldFilled);

                if (allFilled) {
                    link.classList.add('complete');
                } else {
                    link.classList.add('incomplete');
                }
            } else if (section.optionalFields && section.optionalFields.length > 0) {
                // Optional section - check if any fields are filled
                const anyFilled = section.optionalFields.some(isFieldFilled);

                if (anyFilled) {
                    link.classList.add('complete');
                } else {
                    link.classList.add('optional');
                }
            } else {
                // No fields defined, mark as optional
                link.classList.add('optional');
            }
        });
    }

    loadSampleData() {
        const sampleData = {
            courseTitle: "Strategic Marketing Management",
            courseNumber: "MKTG-GB.2334.01",
            term: "Spring 2025",
            credits: "3",
            prerequisites: "Core Marketing (MKTG-GB.2109) or equivalent",
            instructorName: "Dr. Jane Smith",
            officeHours: "Tuesdays and Thursdays, 3:30-5:00 PM, Room 7-65, or by appointment",
            classSchedule: "MW 1:30-2:50 PM, Room KMC 4-80",
            courseDescription: "This course examines strategic marketing decisions facing firms in competitive markets. Students will learn frameworks for analyzing market opportunities, developing positioning strategies, and designing marketing programs. The course emphasizes both analytical rigor and practical application through case studies and simulations.",
            learningOutcomes: "At the conclusion of this course, students will be able to:\n- Analyze competitive market dynamics and identify strategic opportunities\n- Develop effective market segmentation and targeting strategies\n- Create compelling value propositions and positioning strategies\n- Design integrated marketing programs across the 4Ps\n- Measure and optimize marketing performance using key metrics",
            communicationStrategy: "I am available during office hours and by appointment. Please email me at js123@stern.nyu.edu for questions. I typically respond within 24 hours on weekdays. For urgent matters, please note 'URGENT' in the subject line.",
            technicalRequirements: "Students will need access to:\n- NYU Brightspace for course materials and assignments\n- Microsoft Excel or Google Sheets for case analysis\n- Zoom for any virtual sessions\n- Access to Harvard Business Publishing course pack",
            assignmentTypes: "1. Case Analysis (Individual): 3 written case analyses, 3-5 pages each\n2. Group Project: Strategic marketing plan for a real company\n3. Midterm Exam: In-class exam covering frameworks and concepts\n4. Class Participation: Active engagement in case discussions",
            gradingPercentages: "Case Analyses: 30% (10% each)\nGroup Project: 30%\nMidterm Exam: 25%\nClass Participation: 15%",
            dueDatesPolicy: "All assignments are due by 11:59 PM on the specified date via Brightspace. Late submissions will be penalized 10% per day. Extensions may be granted for documented emergencies with advance notice.",
            courseOutline: "Week 1: Introduction to Strategic Marketing\nWeek 2-3: Market Analysis and Segmentation\nWeek 4-5: Positioning and Value Proposition Design\nWeek 6-7: Product and Pricing Strategies\nWeek 8: Midterm Exam\nWeek 9-10: Distribution and Channel Management\nWeek 11-12: Marketing Communications\nWeek 13-14: Digital Marketing and Analytics\nWeek 15: Group Presentations and Course Wrap-up",
            academicIntegrity: "Academic integrity is fundamental to the educational process and all members of the NYU Stern community are expected to act in accordance with the highest level of academic honesty. Violations will not be tolerated and may result in failure of the assignment, failure of the course, suspension or expulsion from the program.",
            integrityOfCredit: "Students will meet 2x a week for 1 hour 20 minutes each session for 15 weeks for this 3-credit course, totaling 40 contact hours.",
            gradingGuidelines: "At NYU Stern, we strive to create courses that challenge students intellectually and that meet the Stern standards of academic excellence. To ensure fairness and clarity of grading, the Stern faculty have adopted a grading guideline for core courses with enrollments of more than 25 students in which approximately 35% of students will receive an \"A\" or \"A-\" grade.",
            studentWellness: "Comprehensive Support: NYU provides a range of support services for student wellness including counseling, health services, and academic support. Students are encouraged to seek help when needed and to prioritize their mental and physical health.",
            religiousObservances: "Accommodating Policy: NYU respects the rights of students to observe religious holidays. Students should notify the instructor in advance of religious observances that may affect class attendance or assignment due dates. Reasonable accommodations will be made.",
            electronicDevices: "Laptops Allowed: Laptops and tablets are permitted for note-taking and course-related activities. Please use devices respectfully and avoid non-course related activities during class.",
            aiGuidance: "Limited AI Use: AI tools may be used for brainstorming and initial research only. All final work must be original, and any AI assistance must be disclosed and properly cited. AI-generated content may not constitute more than 10% of any assignment."
        };

        this.populateForm(sampleData);
        alert('Sample data loaded successfully!');
    }

    // Collect all form data
    collectFormData() {
        const formData = new FormData(this.form);
        const data = {};

        for (let [key, value] of formData.entries()) {
            data[key] = value;
        }

        // Add course outline data
        data.courseOutline = this.collectCourseOutlineData();

        return data;
    }

    // Get the display text from select dropdowns
    getSelectDisplayText(fieldName, value) {
        const element = document.getElementById(fieldName);
        if (element && element.tagName === 'SELECT') {
            const option = element.querySelector(`option[value="${value}"]`);
            return option ? option.text : value;
        }
        return value;
    }

    // Convert text with line breaks to array of Paragraphs
    createParagraphsFromText(text) {
        if (!text) return [];
        const lines = text.split('\n');
        return lines.map(line => new Paragraph({
            children: [new TextRun({ text: line || ' ' })]
        }));
    }

    // Populate form with data
    populateForm(data) {
        Object.keys(data).forEach(key => {
            if (key === 'courseOutline') {
                // Handle course outline separately
                this.populateCourseOutline(data[key]);
            } else {
                const element = document.getElementById(key);
                if (element) {
                    if (element.type === 'checkbox') {
                        element.checked = data[key] === 'true';
                    } else {
                        element.value = data[key];
                    }
                }
            }
        });
        // Update TOC status after populating form
        this.updateTOCStatus();
    }

    populateCourseOutline(outlineData) {
        if (!outlineData || typeof outlineData === 'string') {
            // Legacy format - just text
            return;
        }

        // Clear existing items
        const moduleContainer = document.getElementById('moduleContainer');
        moduleContainer.innerHTML = '';

        // Populate modules with nested class days
        if (outlineData.modules && outlineData.modules.length > 0) {
            outlineData.modules.forEach(module => {
                const moduleId = 'module-' + Date.now() + Math.random();
                const moduleItem = document.createElement('div');
                moduleItem.className = 'module-item';
                moduleItem.dataset.moduleId = moduleId;
                moduleItem.innerHTML = `
                    <div class="module-header">
                        <input type="text" class="module-title" placeholder="Module: Introduction to Topic" value="${module.title || ''}">
                        <button type="button" class="remove-module-btn">Remove Module</button>
                    </div>
                    <textarea class="module-description" rows="2" placeholder="Brief module description or learning objectives (optional)">${module.description || ''}</textarea>
                    <div class="module-class-days" data-module-id="${moduleId}">
                        <!-- Class days for this module -->
                    </div>
                    <button type="button" class="add-class-day-btn" data-module-id="${moduleId}">Add Class Day</button>
                `;
                moduleContainer.appendChild(moduleItem);

                // Populate class days for this module
                const classDaysContainer = moduleItem.querySelector('.module-class-days');
                if (module.classDays && module.classDays.length > 0) {
                    module.classDays.forEach(day => {
                        const classDayItem = document.createElement('div');
                        classDayItem.className = 'class-day-item';
                        classDayItem.innerHTML = `
                            <div class="class-day-header">
                                <input type="text" class="class-day-title" placeholder="Class Day: Topic Name" value="${day.title || ''}" required>
                                <button type="button" class="remove-class-btn">Remove</button>
                            </div>
                            <textarea class="class-day-content" rows="6" placeholder="Readings: ..." required>${day.content || ''}</textarea>
                        `;
                        classDaysContainer.appendChild(classDayItem);
                    });
                }
            });
        } else {
            // Add default module
            this.addModule();
        }
    }

    // Save syllabus as JSON for editing
    saveSyllabus() {
        const data = this.collectFormData();
        const dataStr = JSON.stringify(data, null, 2);
        const blob = new Blob([dataStr], { type: 'application/json' });
        
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `syllabus_${data.courseNumber || 'draft'}_${new Date().toISOString().split('T')[0]}.json`;
        link.click();
    }

    // Load syllabus from JSON file
    loadSyllabus() {
        document.getElementById('loadFile').click();
    }

    handleFileLoad(event) {
        const file = event.target.files[0];
        if (file && file.type === 'application/json') {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = JSON.parse(e.target.result);
                    this.populateForm(data);
                    alert('Syllabus loaded successfully!');
                } catch (error) {
                    alert('Error loading file: Invalid JSON format');
                }
            };
            reader.readAsText(file);
        } else {
            alert('Please select a valid JSON file');
        }
    }

    // Export to Word document
    async exportToWord() {
        const data = this.collectFormData();

        // Get actual text for select fields
        data.academicIntegrity = this.getSelectDisplayText('academicIntegrity', data.academicIntegrity);
        data.gradingGuidelines = this.getSelectDisplayText('gradingGuidelines', data.gradingGuidelines);
        if (data.studentWellness) {
            data.studentWellness = this.getSelectDisplayText('studentWellness', data.studentWellness);
        }
        if (data.religiousObservances) {
            data.religiousObservances = this.getSelectDisplayText('religiousObservances', data.religiousObservances);
        }
        if (data.electronicDevices) {
            data.electronicDevices = this.getSelectDisplayText('electronicDevices', data.electronicDevices);
        }
        if (data.aiGuidance) {
            data.aiGuidance = this.getSelectDisplayText('aiGuidance', data.aiGuidance);
        }

        // Validation
        const requiredFields = ['courseTitle', 'courseNumber', 'term', 'credits', 'instructorName', 'officeHours', 'classSchedule', 'courseDescription', 'learningOutcomes', 'assignmentTypes', 'gradingPercentages', 'dueDatesPolicy', 'academicIntegrity', 'integrityOfCredit', 'gradingGuidelines'];

        const missingFields = requiredFields.filter(field => !data[field] || data[field].trim() === '');

        if (missingFields.length > 0) {
            alert(`Please fill in all required fields: ${missingFields.map(f => f.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase())).join(', ')}`);
            return;
        }

        // Validate course outline
        if (!data.courseOutline || !data.courseOutline.modules || data.courseOutline.modules.length === 0) {
            alert('Please add at least one module to the course outline.');
            return;
        }

        // Check that at least one module has class days
        const hasClassDays = data.courseOutline.modules.some(module =>
            module.classDays && module.classDays.length > 0
        );
        if (!hasClassDays) {
            alert('Please add at least one class day to the course outline.');
            return;
        }

        try {
            // Create Word document using docx library
            const doc = new Document({
                sections: [{
                    properties: {},
                    children: [
                        // Title
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: data.courseTitle,
                                    bold: true,
                                    size: 32
                                })
                            ],
                            alignment: AlignmentType.CENTER
                        }),
                        
                        // Basic course info
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Course Number / Section: ${data.courseNumber}`,
                                    break: 1
                                })
                            ]
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Term: ${data.term}`
                                })
                            ]
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Credits: ${data.credits}`
                                })
                            ]
                        }),
                        
                        ...(data.prerequisites ? [new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Prerequisites: ${data.prerequisites}`
                                })
                            ]
                        })] : []),
                        
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Instructor Name: ${data.instructorName}`
                                })
                            ]
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Office Hours: ${data.officeHours}`
                                })
                            ]
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Class Schedule: ${data.classSchedule}`
                                })
                            ]
                        }),
                        
                        // Course Description
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Course Description",
                                    bold: true,
                                    size: 24,
                                    break: 2
                                })
                            ]
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: data.courseDescription
                                })
                            ]
                        }),
                        
                        // Learning Outcomes
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Learning Outcomes",
                                    bold: true,
                                    size: 24,
                                    break: 2
                                })
                            ]
                        }),
                        ...this.createParagraphsFromText(data.learningOutcomes),
                        
                        // Communication Strategy
                        ...(data.communicationStrategy ? [
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: "Communication Strategy",
                                        bold: true,
                                        size: 24,
                                        break: 2
                                    })
                                ]
                            }),
                            ...this.createParagraphsFromText(data.communicationStrategy)
                        ] : []),
                        
                        // Technical Requirements
                        ...(data.technicalRequirements ? [
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: "Technical Requirements",
                                        bold: true,
                                        size: 24,
                                        break: 2
                                    })
                                ]
                            }),
                            ...this.createParagraphsFromText(data.technicalRequirements)
                        ] : []),
                        
                        // Course Requirements and Assignments
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Course Requirements and Assignments",
                                    bold: true,
                                    size: 24,
                                    break: 2
                                })
                            ]
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Assignment Types and Descriptions:",
                                    bold: true,
                                    break: 1
                                })
                            ]
                        }),
                        ...this.createParagraphsFromText(data.assignmentTypes),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Grading Percentages:",
                                    bold: true,
                                    break: 1
                                })
                            ]
                        }),
                        ...this.createParagraphsFromText(data.gradingPercentages),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Due Dates and Late Policy:",
                                    bold: true,
                                    break: 1
                                })
                            ]
                        }),
                        ...this.createParagraphsFromText(data.dueDatesPolicy),
                        
                        // Course Outline
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Course Outline",
                                    bold: true,
                                    size: 24,
                                    break: 2
                                })
                            ]
                        }),
                        ...this.createParagraphsFromText(this.formatCourseOutlineForExport(data.courseOutline)),
                        
                        // All the required sections
                        ...this.generateRequiredSections(data)
                    ]
                }]
            });

            const blob = await Packer.toBlob(doc);
            saveAs(blob, `${data.courseNumber}_Syllabus.docx`);
            
        } catch (error) {
            console.error('Error creating Word document:', error);
            alert('Error creating Word document. Please try again.');
        }
    }

    generateRequiredSections(data) {
        const sections = [];
        
        // Academic Integrity
        sections.push(
            new Paragraph({
                children: [
                    new TextRun({
                        text: "Academic Integrity",
                        bold: true,
                        size: 24,
                        break: 2
                    })
                ]
            }),
            new Paragraph({
                children: [
                    new TextRun({
                        text: data.academicIntegrity
                    })
                ]
            })
        );
        
        // Stern Code of Conduct
        sections.push(
            new Paragraph({
                children: [
                    new TextRun({
                        text: "Stern Code of Conduct",
                        bold: true,
                        size: 24,
                        break: 2
                    })
                ]
            }),
            ...this.createParagraphsFromText(data.codeOfConduct)
        );
        
        // Integrity of Credit
        sections.push(
            new Paragraph({
                children: [
                    new TextRun({
                        text: "Integrity of Credit",
                        bold: true,
                        size: 24,
                        break: 2
                    })
                ]
            }),
            new Paragraph({
                children: [
                    new TextRun({
                        text: data.integrityOfCredit
                    })
                ]
            })
        );
        
        // General Conduct
        sections.push(
            new Paragraph({
                children: [
                    new TextRun({
                        text: "General Conduct & Behavior",
                        bold: true,
                        size: 24,
                        break: 2
                    })
                ]
            }),
            ...this.createParagraphsFromText(data.generalConduct)
        );
        
        // Grading Guidelines
        if (data.gradingGuidelines) {
            sections.push(
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Grading Guidelines",
                            bold: true,
                            size: 24,
                            break: 2
                        })
                    ]
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: data.gradingGuidelines
                        })
                    ]
                })
            );
        }
        
        // Student Accessibility
        sections.push(
            new Paragraph({
                children: [
                    new TextRun({
                        text: "Student Accessibility",
                        bold: true,
                        size: 24,
                        break: 2
                    })
                ]
            }),
            ...this.createParagraphsFromText(data.studentAccessibility)
        );
        
        // Optional sections
        if (data.studentWellness) {
            sections.push(
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Student Wellness",
                            bold: true,
                            size: 24,
                            break: 2
                        })
                    ]
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: data.studentWellness
                        })
                    ]
                })
            );
        }
        
        if (data.namePronouns) {
            sections.push(
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Name Pronunciation and Pronouns",
                            bold: true,
                            size: 24,
                            break: 2
                        })
                    ]
                }),
                ...this.createParagraphsFromText(data.namePronouns)
            );
        }
        
        if (data.religiousObservances) {
            sections.push(
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Religious Observances and Absences",
                            bold: true,
                            size: 24,
                            break: 2
                        })
                    ]
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: data.religiousObservances
                        })
                    ]
                })
            );
        }
        
        if (data.electronicDevices) {
            sections.push(
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Electronic Devices Policy",
                            bold: true,
                            size: 24,
                            break: 2
                        })
                    ]
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: data.electronicDevices
                        })
                    ]
                })
            );
        }
        
        if (data.aiGuidance) {
            sections.push(
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "AI Guidance",
                            bold: true,
                            size: 24,
                            break: 2
                        })
                    ]
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: data.aiGuidance
                        })
                    ]
                })
            );
        }
        
        return sections;
    }

    // Export to PDF
    async exportToPDF() {
        const data = this.collectFormData();

        // Get actual text for select fields
        data.academicIntegrity = this.getSelectDisplayText('academicIntegrity', data.academicIntegrity);
        data.gradingGuidelines = this.getSelectDisplayText('gradingGuidelines', data.gradingGuidelines);
        if (data.studentWellness) {
            data.studentWellness = this.getSelectDisplayText('studentWellness', data.studentWellness);
        }
        if (data.religiousObservances) {
            data.religiousObservances = this.getSelectDisplayText('religiousObservances', data.religiousObservances);
        }
        if (data.electronicDevices) {
            data.electronicDevices = this.getSelectDisplayText('electronicDevices', data.electronicDevices);
        }
        if (data.aiGuidance) {
            data.aiGuidance = this.getSelectDisplayText('aiGuidance', data.aiGuidance);
        }

        // Validation
        const requiredFields = ['courseTitle', 'courseNumber', 'term', 'credits', 'instructorName', 'officeHours', 'classSchedule', 'courseDescription', 'learningOutcomes', 'assignmentTypes', 'gradingPercentages', 'dueDatesPolicy', 'academicIntegrity', 'integrityOfCredit', 'gradingGuidelines'];

        const missingFields = requiredFields.filter(field => !data[field] || data[field].trim() === '');

        if (missingFields.length > 0) {
            alert(`Please fill in all required fields: ${missingFields.map(f => f.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase())).join(', ')}`);
            return;
        }

        // Validate course outline
        if (!data.courseOutline || !data.courseOutline.modules || data.courseOutline.modules.length === 0) {
            alert('Please add at least one module to the course outline.');
            return;
        }

        // Check that at least one module has class days
        const hasClassDays = data.courseOutline.modules.some(module =>
            module.classDays && module.classDays.length > 0
        );
        if (!hasClassDays) {
            alert('Please add at least one class day to the course outline.');
            return;
        }

        try {
            const { jsPDF } = window.jspdf;
            const pdf = new jsPDF('p', 'mm', 'a4');

            const pageWidth = 210;
            const pageHeight = 297;
            const margin = 20;
            const contentWidth = pageWidth - (2 * margin);
            let yPosition = margin;

            // Helper function to add text with auto page breaks
            const addText = (text, fontSize, isBold = false, color = [0, 0, 0]) => {
                pdf.setFontSize(fontSize);
                pdf.setFont('helvetica', isBold ? 'bold' : 'normal');
                pdf.setTextColor(...color);

                const lines = pdf.splitTextToSize(text, contentWidth);

                for (let line of lines) {
                    if (yPosition + fontSize / 2 > pageHeight - margin) {
                        pdf.addPage();
                        yPosition = margin;
                    }
                    pdf.text(line, margin, yPosition);
                    yPosition += fontSize / 2 + 2;
                }
            };

            const addSpacing = (space = 5) => {
                yPosition += space;
            };

            // Title
            pdf.setTextColor(87, 6, 140); // NYU Purple
            addText(data.courseTitle, 18, true, [87, 6, 140]);
            addText('NYU Stern School of Business', 14, false, [87, 6, 140]);
            addSpacing(10);

            // Basic Info
            pdf.setTextColor(0, 0, 0);
            addText(`Course Number / Section: ${data.courseNumber}`, 11, true);
            addText(`Term: ${data.term}`, 11);
            addText(`Credits: ${data.credits}`, 11);
            if (data.prerequisites) addText(`Prerequisites: ${data.prerequisites}`, 11);
            addText(`Instructor Name: ${data.instructorName}`, 11);
            addText(`Office Hours: ${data.officeHours}`, 11);
            addText(`Class Schedule: ${data.classSchedule}`, 11);
            addSpacing(10);

            // Course Description
            addText('Course Description', 14, true, [87, 6, 140]);
            addSpacing(3);
            addText(data.courseDescription, 11);
            addSpacing(10);

            // Learning Outcomes
            addText('Learning Outcomes', 14, true, [87, 6, 140]);
            addSpacing(3);
            addText(data.learningOutcomes, 11);
            addSpacing(10);

            // Communication Strategy
            if (data.communicationStrategy) {
                addText('Communication Strategy', 14, true, [87, 6, 140]);
                addSpacing(3);
                addText(data.communicationStrategy, 11);
                addSpacing(10);
            }

            // Technical Requirements
            if (data.technicalRequirements) {
                addText('Technical Requirements', 14, true, [87, 6, 140]);
                addSpacing(3);
                addText(data.technicalRequirements, 11);
                addSpacing(10);
            }

            // Course Requirements and Assignments
            addText('Course Requirements and Assignments', 14, true, [87, 6, 140]);
            addSpacing(3);
            addText('Assignment Types and Descriptions:', 11, true);
            addText(data.assignmentTypes, 11);
            addSpacing(5);
            addText('Grading Percentages:', 11, true);
            addText(data.gradingPercentages, 11);
            addSpacing(5);
            addText('Due Dates and Late Policy:', 11, true);
            addText(data.dueDatesPolicy, 11);
            addSpacing(10);

            // Course Outline
            addText('Course Outline', 14, true, [87, 6, 140]);
            addSpacing(3);
            addText(this.formatCourseOutlineForExport(data.courseOutline), 11);
            addSpacing(10);

            // Academic Integrity
            addText('Academic Integrity', 14, true, [87, 6, 140]);
            addSpacing(3);
            addText(data.academicIntegrity, 11);
            addSpacing(10);

            // Stern Code of Conduct
            addText('Stern Code of Conduct', 14, true, [87, 6, 140]);
            addSpacing(3);
            addText(data.codeOfConduct, 11);
            addSpacing(10);

            // Integrity of Credit
            addText('Integrity of Credit', 14, true, [87, 6, 140]);
            addSpacing(3);
            addText(data.integrityOfCredit, 11);
            addSpacing(10);

            // General Conduct
            addText('General Conduct & Behavior', 14, true, [87, 6, 140]);
            addSpacing(3);
            addText(data.generalConduct, 11);
            addSpacing(10);

            // Grading Guidelines
            if (data.gradingGuidelines) {
                addText('Grading Guidelines', 14, true, [87, 6, 140]);
                addSpacing(3);
                addText(data.gradingGuidelines, 11);
                addSpacing(10);
            }

            // Student Accessibility
            addText('Student Accessibility', 14, true, [87, 6, 140]);
            addSpacing(3);
            addText(data.studentAccessibility, 11);
            addSpacing(10);

            // Optional sections
            if (data.studentWellness) {
                addText('Student Wellness', 14, true, [87, 6, 140]);
                addSpacing(3);
                addText(data.studentWellness, 11);
                addSpacing(10);
            }

            if (data.namePronouns) {
                addText('Name Pronunciation and Pronouns', 14, true, [87, 6, 140]);
                addSpacing(3);
                addText(data.namePronouns, 11);
                addSpacing(10);
            }

            if (data.religiousObservances) {
                addText('Religious Observances and Absences', 14, true, [87, 6, 140]);
                addSpacing(3);
                addText(data.religiousObservances, 11);
                addSpacing(10);
            }

            if (data.electronicDevices) {
                addText('Electronic Devices Policy', 14, true, [87, 6, 140]);
                addSpacing(3);
                addText(data.electronicDevices, 11);
                addSpacing(10);
            }

            if (data.aiGuidance) {
                addText('AI Guidance', 14, true, [87, 6, 140]);
                addSpacing(3);
                addText(data.aiGuidance, 11);
                addSpacing(10);
            }

            pdf.save(`${data.courseNumber}_Syllabus.pdf`);

        } catch (error) {
            console.error('Error creating PDF:', error);
            alert('Error creating PDF. Please try again.');
        }
    }

    generateSyllabusHTML(data) {
        return `
            <div style="text-align: center; margin-bottom: 30px;">
                <h1 style="font-size: 24px; margin: 0; color: #57068C;">${data.courseTitle}</h1>
                <h2 style="font-size: 16px; margin: 10px 0; color: #57068C;">NYU Stern School of Business</h2>
            </div>
            
            <div style="margin-bottom: 25px;">
                <p><strong>Course Number / Section:</strong> ${data.courseNumber}</p>
                <p><strong>Term:</strong> ${data.term}</p>
                <p><strong>Credits:</strong> ${data.credits}</p>
                ${data.prerequisites ? `<p><strong>Prerequisites:</strong> ${data.prerequisites}</p>` : ''}
                <p><strong>Instructor Name:</strong> ${data.instructorName}</p>
                <p><strong>Office Hours:</strong> ${data.officeHours}</p>
                <p><strong>Class Schedule:</strong> ${data.classSchedule}</p>
            </div>
            
            <div style="margin-bottom: 25px;">
                <h3 style="color: #57068C; border-bottom: 2px solid #57068C; padding-bottom: 5px;">Course Description</h3>
                <p>${data.courseDescription}</p>
            </div>
            
            <div style="margin-bottom: 25px;">
                <h3 style="color: #57068C; border-bottom: 2px solid #57068C; padding-bottom: 5px;">Learning Outcomes</h3>
                <div style="white-space: pre-line;">${data.learningOutcomes}</div>
            </div>
            
            ${data.communicationStrategy ? `
            <div style="margin-bottom: 25px;">
                <h3 style="color: #57068C; border-bottom: 2px solid #57068C; padding-bottom: 5px;">Communication Strategy</h3>
                <p>${data.communicationStrategy}</p>
            </div>
            ` : ''}
            
            ${data.technicalRequirements ? `
            <div style="margin-bottom: 25px;">
                <h3 style="color: #57068C; border-bottom: 2px solid #57068C; padding-bottom: 5px;">Technical Requirements</h3>
                <p>${data.technicalRequirements}</p>
            </div>
            ` : ''}
            
            <div style="margin-bottom: 25px;">
                <h3 style="color: #57068C; border-bottom: 2px solid #57068C; padding-bottom: 5px;">Course Requirements and Assignments</h3>
                <p><strong>Assignment Types and Descriptions:</strong></p>
                <div style="white-space: pre-line;">${data.assignmentTypes}</div>
                <p><strong>Grading Percentages:</strong></p>
                <div style="white-space: pre-line;">${data.gradingPercentages}</div>
                <p><strong>Due Dates and Late Policy:</strong></p>
                <div style="white-space: pre-line;">${data.dueDatesPolicy}</div>
            </div>
            
            <div style="margin-bottom: 25px;">
                <h3 style="color: #57068C; border-bottom: 2px solid #57068C; padding-bottom: 5px;">Course Outline</h3>
                ${this.generateCourseOutlineHTML(data.courseOutline)}
            </div>
            
            <div style="margin-bottom: 25px;">
                <h3 style="color: #57068C; border-bottom: 2px solid #57068C; padding-bottom: 5px;">Academic Integrity</h3>
                <p>${data.academicIntegrity}</p>
            </div>
            
            <div style="margin-bottom: 25px;">
                <h3 style="color: #57068C; border-bottom: 2px solid #57068C; padding-bottom: 5px;">Stern Code of Conduct</h3>
                <div style="white-space: pre-line;">${data.codeOfConduct}</div>
            </div>
            
            <div style="margin-bottom: 25px;">
                <h3 style="color: #57068C; border-bottom: 2px solid #57068C; padding-bottom: 5px;">Integrity of Credit</h3>
                <p>${data.integrityOfCredit}</p>
            </div>
            
            <div style="margin-bottom: 25px;">
                <h3 style="color: #57068C; border-bottom: 2px solid #57068C; padding-bottom: 5px;">General Conduct & Behavior</h3>
                <div style="white-space: pre-line;">${data.generalConduct}</div>
            </div>
            
            ${data.gradingGuidelines ? `
            <div style="margin-bottom: 25px;">
                <h3 style="color: #57068C; border-bottom: 2px solid #57068C; padding-bottom: 5px;">Grading Guidelines</h3>
                <p>${data.gradingGuidelines}</p>
            </div>
            ` : ''}
            
            <div style="margin-bottom: 25px;">
                <h3 style="color: #57068C; border-bottom: 2px solid #57068C; padding-bottom: 5px;">Student Accessibility</h3>
                <div style="white-space: pre-line;">${data.studentAccessibility}</div>
            </div>
            
            ${data.studentWellness ? `
            <div style="margin-bottom: 25px;">
                <h3 style="color: #57068C; border-bottom: 2px solid #57068C; padding-bottom: 5px;">Student Wellness</h3>
                <p>${data.studentWellness}</p>
            </div>
            ` : ''}
            
            ${data.namePronouns ? `
            <div style="margin-bottom: 25px;">
                <h3 style="color: #57068C; border-bottom: 2px solid #57068C; padding-bottom: 5px;">Name Pronunciation and Pronouns</h3>
                <div style="white-space: pre-line;">${data.namePronouns}</div>
            </div>
            ` : ''}
            
            ${data.religiousObservances ? `
            <div style="margin-bottom: 25px;">
                <h3 style="color: #57068C; border-bottom: 2px solid #57068C; padding-bottom: 5px;">Religious Observances and Absences</h3>
                <p>${data.religiousObservances}</p>
            </div>
            ` : ''}
            
            ${data.electronicDevices ? `
            <div style="margin-bottom: 25px;">
                <h3 style="color: #57068C; border-bottom: 2px solid #57068C; padding-bottom: 5px;">Electronic Devices Policy</h3>
                <p>${data.electronicDevices}</p>
            </div>
            ` : ''}
            
            ${data.aiGuidance ? `
            <div style="margin-bottom: 25px;">
                <h3 style="color: #57068C; border-bottom: 2px solid #57068C; padding-bottom: 5px;">AI Guidance</h3>  
                <p>${data.aiGuidance}</p>
            </div>
            ` : ''}
        `;
    }
}

// Initialize the app when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    new SyllabusBuilder();
});