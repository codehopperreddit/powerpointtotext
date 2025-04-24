     // DOM elements
     const uploadArea = document.getElementById('uploadArea');
     const fileInput = document.getElementById('fileInput');
     const convertBtn = document.getElementById('convertBtn');
     const downloadBtn = document.getElementById('downloadBtn');
     const textPreview = document.getElementById('textPreview');
     const statusMessage = document.getElementById('statusMessage');
     
     // Global variables
     let selectedFile = null;
     let extractedText = '';
     
     // Event listeners
     uploadArea.addEventListener('click', () => fileInput.click());
     uploadArea.addEventListener('dragover', (e) => {
         e.preventDefault();
         uploadArea.style.backgroundColor = '#f0f0f0';
     });
     uploadArea.addEventListener('dragleave', () => {
         uploadArea.style.backgroundColor = '';
     });
     uploadArea.addEventListener('drop', (e) => {
         e.preventDefault();
         uploadArea.style.backgroundColor = '';
         
         if (e.dataTransfer.files.length) {
             handleFileSelection(e.dataTransfer.files[0]);
         }
     });
     
     fileInput.addEventListener('change', (e) => {
         if (e.target.files.length) {
             handleFileSelection(e.target.files[0]);
         }
     });
     
     convertBtn.addEventListener('click', convertToText);
     downloadBtn.addEventListener('click', downloadText);
     
     // Functions
     function handleFileSelection(file) {
         // Check if the file is a PowerPoint file
         if (!file.name.endsWith('.pptx')) {
             showStatus('Please select a .pptx file (newer PowerPoint format)', 'error');
             return;
         }
         
         selectedFile = file;
         uploadArea.innerHTML = `<p>Selected file: ${file.name}</p>`;
         convertBtn.disabled = false;
         textPreview.textContent = '';
         downloadBtn.disabled = true;
         showStatus('File selected. Click "Convert to Text" to proceed.', 'success');
     }
     
     function convertToText() {
         if (!selectedFile) return;
         
         showStatus('<span class="loading"></span> Converting, please wait...', 'info');
         textPreview.textContent = 'Processing...';
         
         const reader = new FileReader();
         reader.onload = function(e) {
             try {
                 const data = e.target.result;
                 
                 // Use JSZip to extract the PowerPoint content
                 extractPptxText(data).then(text => {
                     extractedText = text;
                     
                     // Display the extracted text
                     textPreview.textContent = extractedText;
                     
                     // Enable download button
                     downloadBtn.disabled = false;
                     
                     showStatus('Conversion completed successfully!', 'success');
                 }).catch(error => {
                     console.error('Error extracting text:', error);
                     showStatus('Error extracting text from PPTX. ' + error.message, 'error');
                     textPreview.textContent = 'Error processing file.';
                 });
             } catch (error) {
                 console.error('Error converting file:', error);
                 showStatus('Error converting file. ' + error.message, 'error');
                 textPreview.textContent = 'Error processing file.';
             }
         };
         
         reader.onerror = function() {
             showStatus('Error reading file', 'error');
             textPreview.textContent = 'Error reading file.';
         };
         
         reader.readAsArrayBuffer(selectedFile);
     }
     
     async function extractPptxText(arrayBuffer) {
         // PowerPoint files are basically ZIP files with XML content
         const zip = new JSZip();
         const contents = await zip.loadAsync(arrayBuffer);
         
         // Extract slides content - look in the ppt/slides directory
         const slideFiles = [];
         const slidePattern = /^ppt\/slides\/slide\d+\.xml$/;
         
         // Collect all slide files
         Object.keys(contents.files).forEach(filename => {
             if (slidePattern.test(filename)) {
                 slideFiles.push({
                     name: filename,
                     index: parseInt(filename.match(/slide(\d+)\.xml/)[1])
                 });
             }
         });
         
         // Sort slides by index
         slideFiles.sort((a, b) => a.index - b.index);
         
         let allText = '';
         
         // Process each slide
         for (const slideFile of slideFiles) {
             const slideContent = await contents.file(slideFile.name).async('string');
             const slideText = extractTextFromSlideXML(slideContent);
             
             if (slideText.trim()) {
                 allText += `===== Slide ${slideFile.index} =====\n${slideText}\n\n`;
             }
         }
         
         // Look for notes if available
         const notesPattern = /^ppt\/notesSlides\/notesSlide\d+\.xml$/;
         const notesFiles = [];
         
         Object.keys(contents.files).forEach(filename => {
             if (notesPattern.test(filename)) {
                 notesFiles.push({
                     name: filename,
                     index: parseInt(filename.match(/notesSlide(\d+)\.xml/)[1])
                 });
             }
         });
         
         // Sort notes by index
         notesFiles.sort((a, b) => a.index - b.index);
         
         // Process each notes slide
         if (notesFiles.length > 0) {
             allText += "\n===== SLIDE NOTES =====\n\n";
             
             for (const notesFile of notesFiles) {
                 const notesContent = await contents.file(notesFile.name).async('string');
                 const notesText = extractTextFromSlideXML(notesContent);
                 
                 if (notesText.trim()) {
                     allText += `Notes for Slide ${notesFile.index}:\n${notesText}\n\n`;
                 }
             }
         }
         
         return allText || 'No text found in the PowerPoint file.';
     }
     
     function extractTextFromSlideXML(xmlString) {
         const parser = new DOMParser();
         const xmlDoc = parser.parseFromString(xmlString, "text/xml");
         
         // Find all text elements - this will vary based on PowerPoint XML structure
         // Most text in PowerPoint is within <a:t> tags
         const textNodes = xmlDoc.getElementsByTagName('a:t');
         let slideText = '';
         
         for (let i = 0; i < textNodes.length; i++) {
             const textContent = textNodes[i].textContent.trim();
             if (textContent) {
                 slideText += textContent + '\n';
             }
         }
         
         return slideText;
     }
     
     function downloadText() {
         if (!extractedText) return;
         
         // Create a blob with the text
         const blob = new Blob([extractedText], {type: 'text/plain'});
         
         // Create a download link
         const url = URL.createObjectURL(blob);
         const a = document.createElement('a');
         a.href = url;
         a.download = selectedFile.name.replace(/\.pptx$/i, '.txt');
         document.body.appendChild(a);
         a.click();
         
         // Clean up
         setTimeout(() => {
             document.body.removeChild(a);
             URL.revokeObjectURL(url);
         }, 0);
         
         showStatus('Text file downloaded successfully!', 'success');
     }
     
     function showStatus(message, type) {
         statusMessage.innerHTML = message;
         statusMessage.className = 'status ' + type;
         statusMessage.style.display = 'block';
         
         // Hide the message after 5 seconds if it's a success message
         if (type === 'success') {
             setTimeout(() => {
                 statusMessage.style.display = 'none';
             }, 5000);
         }
     }
 </script>