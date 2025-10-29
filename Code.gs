/**
 * ═══════════════════════════════════════════════════════════════
 * NAGWA LATEX CONVERTER - COMPLETE VERSION WITH HANDWRITING
 * Google Slides to LaTeX Beamer Converter
 *
 * Features:
 * - Auto-detect annotations from shapes (circles, rectangles)
 * - Auto-detect from text colors (red=circle, blue=box, green=underline)
 * - Auto-detect from speaker notes (Word: explanation format)
 * - Manual annotation sidebar
 * - Visual handwriting placement via auto-created placeholder
 * - Export to Nagwa LaTeX format
 * - Download .tex file to Google Drive
 *
 * Version: 2.1 (Added Auto-Placeholder Creation)
 * ═══════════════════════════════════════════════════════════════
 */

// ══════════════════════════════════════════════════════
// CONFIGURATION
// ══════════════════════════════════════════════════════

const DEFAULT_CONFIG = {
  sessionID: '000000000000',
  country: 'eg',
  subject: 'Physics',
  language: 'en',
  grade: '12',
  term: '0',
  sessionTitle: 'Untitled Session',
  documentClass: 'Worksheet',
  numerals: 'european',
  directions: 'ltr'
};

// Handwriting configuration
const HANDWRITE_MARKER = '_IS_HANDWRITE_ID_=';
const HANDWRITE_BG_COLOR = '#FFFACD'; // لون أصفر فاتح لتمييز المربع

// Coordinate conversion
const SLIDE_WIDTH_EMU = 9144000;
const SLIDE_HEIGHT_EMU = 6858000;
const BEAMER_WIDTH_CM = 12.8;
const BEAMER_HEIGHT_CM = 9.6;

// ══════════════════════════════════════════════════════
// MENU
// ══════════════════════════════════════════════════════

function onOpen() {
  SlidesApp.getUi()
    .createMenu('📚 Nagwa LaTeX')
    .addItem('🔧 Configure Session', 'showConfigDialog')
    .addSeparator()
    .addSubMenu(SlidesApp.getUi().createMenu('📝 Annotations')
      .addItem('🎨 Auto-Detect from Colors', 'autoDetectFromColors')
      .addItem('📐 Auto-Detect from Shapes', 'autoDetectFromShapes')
      .addItem('🖊️ Manual Entry (Sidebar)', 'showAnnotationSidebar')
      .addSeparator()
      .addItem('👁️ Preview All Annotations', 'previewCapturedData')
      .addItem('🗑️ Clear Auto-Detected', 'clearCapturedData'))
    .addSeparator()
    .addSubMenu(SlidesApp.getUi().createMenu('✍️ Handwriting')
      .addItem('🎯 Open Handwriting Panel', 'showHandwritingSidebar') // Sidebar now uses addHandwritingPlaceholder
      .addItem('📋 List All Handwriting', 'listAllHandwriting')
      .addItem('🗑️ Clear All Handwriting', 'clearAllHandwriting'))
    .addSeparator()
    .addItem('🚀 Export to LaTeX', 'exportToNagwaLatex')
    .addItem('📥 Download .tex File', 'downloadTexFile')
    .addSeparator()
    .addItem('❓ Help', 'showHelp')
    .addToUi();
}

// ══════════════════════════════════════════════════════
// CONFIG MANAGEMENT
// ══════════════════════════════════════════════════════

function getSessionConfig() {
  const props = PropertiesService.getDocumentProperties();
  const config = props.getProperty('nagwa_config');

  if (config) {
    return JSON.parse(config);
  }

  const presentation = SlidesApp.getActivePresentation();
  const defaultConfig = Object.assign({}, DEFAULT_CONFIG);
  defaultConfig.sessionTitle = presentation.getName();

  return defaultConfig;
}

function saveSessionConfig(config) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('nagwa_config', JSON.stringify(config));
  return { success: true };
}

function showConfigDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ConfigDialog')
    .setWidth(500)
    .setHeight(650);
  SlidesApp.getUi().showModalDialog(html, '🔧 Session Configuration');
}

// ══════════════════════════════════════════════════════
// HANDWRITING MANAGEMENT (Using Auto-Placeholder Creation)
// ══════════════════════════════════════════════════════

function showHandwritingSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('HandwritingSidebar')
    .setTitle('✍️ Handwriting Manager')
    .setWidth(350);

  SlidesApp.getUi().showSidebar(html);
}

/**
 * @description تنشئ مربع نص جديدًا على الشريحة الحالية ليكون بمثابة علامة بصرية للكتابة اليدوية،
 * وتربطه بالبيانات المدخلة (النص والحجم) عبر معرف فريد يُخزن في Alt Text و Speaker Notes.
 * **هذه الدالة تستدعى من الشريط الجانبي.**
 *
 * @param {string} text النص الفعلي للكتابة اليدوية المراد إضافته.
 * @param {string|number} size حجم الخط المطلوب (كمثال: '12').
 * @returns {object} كائن يحتوي على حالة النجاح والمعرف الفريد الذي تم إنشاؤه.
 * @throws {Error} إذا لم يتم تحديد شريحة أو حدث خطأ.
 */
function addHandwritingPlaceholder(text, size) {
  // التحقق من المدخلات
  if (!text || !size) {
    throw new Error('النص وحجم الخط مطلوبان.');
  }
  const numericSize = parseInt(size);
  if (isNaN(numericSize) || numericSize < 8 || numericSize > 72) {
     throw new Error('حجم الخط غير صالح. يجب أن يكون رقمًا بين 8 و 72.');
  }

  // الحصول على الشريحة الحالية
  const presentation = SlidesApp.getActivePresentation();
  let slide = presentation.getSelection()?.getCurrentPage()?.asSlide();

  // إذا لم تكن هناك شريحة محددة، استخدم الشريحة الأولى كبديل
  if (!slide) {
    const slides = presentation.getSlides();
    if (slides.length > 0) {
       slide = slides[0];
       Logger.log('لم يتم تحديد شريحة، تم استخدام الشريحة الأولى.');
    } else {
       throw new Error('لا توجد شرائح في العرض التقديمي.');
    }
  }

  // --- إنشاء مربع النص الحاضن ---

  // تحديد أبعاد وموقع افتراضي (منتصف الشريحة)
  const slideWidth = presentation.getPageWidth();
  const slideHeight = presentation.getPageHeight();
  const boxWidth = 200; // عرض افتراضي مناسب
  const boxHeight = 50; // ارتفاع افتراضي مناسب
  const left = (slideWidth / 2) - (boxWidth / 2);
  const top = (slideHeight / 2) - (boxHeight / 2);

  // إنشاء الشكل (مربع النص)
  const shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left, top, boxWidth, boxHeight);
  Logger.log('تم إنشاء مربع النص الحاضن بنجاح. ID: ' + shape.getObjectId());

  // --- ربط البيانات وتخزينها ---

  // توليد معرف فريد
  const timestamp = new Date().getTime();
  const handwriteId = 'hw-' + timestamp + '-' + Math.random().toString(36).substring(2, 7);

  // تخزين المعرف الفريد في Alt Text
  const altText = HANDWRITE_MARKER + handwriteId;
  shape.setDescription(altText);
  Logger.log('تم تعيين Alt Text: ' + altText);

  // تخزين البيانات (المعرف، الحجم، النص) في Speaker Notes
  const notesShape = slide.getNotesPage().getSpeakerNotesShape();
  if (!notesShape) {
     Logger.log('لا يمكن الوصول إلى Speaker Notes لهذه الشريحة.');
     // قد نحتاج لإظهار تحذير للمستخدم هنا
  } else {
    const existingNotes = notesShape.getText().asString();
    const handwriteLine = '[handwrite:' + handwriteId + ':' + numericSize + ':' + text + ']';

    let updatedNotes;
    if (existingNotes.trim().length === 0) {
      updatedNotes = handwriteLine;
    } else {
      updatedNotes = existingNotes + '\n' + handwriteLine;
    }
    notesShape.getText().setText(updatedNotes);
    Logger.log('تم تحديث Speaker Notes.');
  }

  // --- (اختياري) تنسيق بصري للمربع الحاضن ---
  try {
    const previewText = 'HW: ' + (text.length > 20 ? text.substring(0, 17) + '...' : text);
    shape.getText().setText(previewText);
    shape.getText().getTextStyle().setForegroundColor('#5F6368'); // Gray for preview

    const fill = shape.getFill();
    fill.setSolidFill(HANDWRITE_BG_COLOR); // Light yellow background

    // Optional: Add a subtle border
    // shape.getBorder().setDashStyle(SlidesApp.DashStyle.SOLID).setWeight(1).getLineFill().setSolidFill('#FDB813');

  } catch (e) {
    Logger.log('خطأ أثناء تنسيق المربع الحاضن (تجاهله): ' + e);
  }

  // جعل الشكل المحدد للمستخدم ليسهل تحريكه فورًا
  shape.select();

  return {
    success: true,
    id: handwriteId,
    message: 'تم إنشاء علامة الكتابة اليدوية! حركها وضعها في المكان الصحيح.'
  };
}


function getSelectedElementInfo() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();

  if (!selection) {
    return { hasSelection: false };
  }

  const selectionType = selection.getSelectionType();

  if (selectionType !== SlidesApp.SelectionType.PAGE_ELEMENT) {
    return { hasSelection: false };
  }

  const elements = selection.getPageElementRange().getPageElements();

  if (elements.length === 0) {
    return { hasSelection: false };
  }

  // Check only the first selected element
  const element = elements[0];
  const elementType = element.getPageElementType();

  if (elementType !== SlidesApp.PageElementType.SHAPE || !element.asShape().getText) {
    // It's not a shape or not a text box shape
    return {
      hasSelection: true,
      isValid: false,
      message: 'Selected element is not a text box/shape.'
    };
  }

  const shape = element.asShape();
  const description = shape.getDescription();

  let isHandwriting = false;
  let handwriteId = null;
  let currentText = shape.getText().asString(); // Get the current preview text

  if (description && description.startsWith(HANDWRITE_MARKER)) {
    isHandwriting = true;
    handwriteId = description.substring(HANDWRITE_MARKER.length);
  }

  return {
    hasSelection: true,
    isValid: true,
    elementId: element.getObjectId(),
    isHandwriting: isHandwriting,
    handwriteId: handwriteId,
    currentText: currentText, // This is the preview text inside the box
    message: isHandwriting ?
      'This box represents handwriting (ID: ' + handwriteId + ')' :
      'Selected text box is ready to be assigned as handwriting.'
  };
}


function getCurrentSlideHandwriting() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();

  let slide;

  if (selection && selection.getSelectionType() === SlidesApp.SelectionType.PAGE) {
    slide = selection.getCurrentPage().asSlide();
  } else {
    // Fallback to the first slide if no page is selected
    const slides = presentation.getSlides();
    if (slides.length > 0) {
        slide = slides[0];
    } else {
        return []; // No slides
    }
  }

  const handwritingList = [];
  const shapes = slide.getShapes();
  const notes = slide.getNotesPage().getSpeakerNotesShape()?.getText()?.asString() || ''; // Handle potential nulls

  shapes.forEach(function(shape) {
    try { // Add try-catch for robustness
      const description = shape.getDescription();

      if (description && description.startsWith(HANDWRITE_MARKER)) {
        const handwriteId = description.substring(HANDWRITE_MARKER.length);

        // Extract data from notes using RegExp
        const regex = new RegExp('\\[handwrite:' + handwriteId.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&') + ':(\\d+):([^\\]]+)\\]'); // Escape ID for regex
        const match = notes.match(regex);

        if (match) {
          handwritingList.push({
            id: handwriteId,
            size: match[1],
            text: match[2],
            shapeId: shape.getObjectId() // Include shape ID for potential future use
          });
        } else {
          Logger.log('Handwriting note data not found for ID: ' + handwriteId);
          // Optionally add placeholder data or skip
        }
      }
    } catch (e) {
      Logger.log('Error processing shape for handwriting: ' + e);
    }
  });

  return handwritingList;
}

function deleteHandwritingById(handwriteId) {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();

  let found = false;
  let errorMsg = null;

  slides.forEach(function(slide) {
    if (found) return; // Stop searching if already found and deleted

    const shapes = slide.getShapes();
    const notesShape = slide.getNotesPage().getSpeakerNotesShape();
    let notesModified = false;
    let notes = notesShape ? notesShape.getText().asString() : '';

    shapes.forEach(function(shape) {
       try {
        const description = shape.getDescription();

        if (description && description === HANDWRITE_MARKER + handwriteId) {
          // 1. Clear shape properties
          shape.setDescription('');
          shape.getText().setText(''); // Clear preview text
          const fill = shape.getFill();
          fill.setSolidFill('#FFFFFF', 0); // Make transparent
          shape.getBorder().getLineFill().setSolidFill('#FFFFFF', 0); // Hide border too
          // Instead of deleting, we make it invisible or clear it. Deleting might cause issues if other elements are grouped.

          // 2. Remove from notes
          if (notesShape) {
            const regex = new RegExp('\\[handwrite:' + handwriteId.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&') + ':[^\\]]+\\]\\n?', 'g');
            const updatedNotes = notes.replace(regex, '');
            if (updatedNotes !== notes) {
               notes = updatedNotes; // Update notes content for potential multiple deletions in same slide
               notesModified = true;
            }
          }
          found = true;
          Logger.log('Handwriting element cleared for ID: ' + handwriteId);
          // Do not break here, continue checking notes in case the line appears multiple times
        }
      } catch (e) {
         Logger.log('Error while deleting handwriting shape properties: ' + e);
         errorMsg = e.message; // Store error message
      }
    });

    // Update notes for the slide if modified
    if (notesModified && notesShape) {
      notesShape.getText().setText(notes);
    }
  });

  if (!found) {
    throw new Error('Handwriting with ID ' + handwriteId + ' not found.');
  }
  if (errorMsg){
     Logger.log('Completed deletion with errors: ' + errorMsg);
     // Decide if you want to throw an error here or just log it
  }

  return { success: true };
}


function listAllHandwriting() {
  const ui = SlidesApp.getUi();
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();

  let list = '';
  let total = 0;

  slides.forEach(function(slide, slideIndex) {
    const shapes = slide.getShapes();
    const notes = slide.getNotesPage().getSpeakerNotesShape()?.getText()?.asString() || '';
    let slideHandwriting = [];

    shapes.forEach(function(shape) {
      try {
        const description = shape.getDescription();
        if (description && description.startsWith(HANDWRITE_MARKER)) {
          const handwriteId = description.substring(HANDWRITE_MARKER.length);
          const regex = new RegExp('\\[handwrite:' + handwriteId.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&') + ':(\\d+):([^\\]]+)\\]');
          const match = notes.match(regex);
          if (match) {
            slideHandwriting.push({
              id: handwriteId,
              size: match[1],
              text: match[2]
            });
          }
        }
      } catch (e) {
          Logger.log('Error listing handwriting for shape on slide ' + (slideIndex + 1) + ': ' + e);
      }
    });

    if (slideHandwriting.length > 0) {
      list += '\n📄 Slide ' + (slideIndex + 1) + ':\n';
      slideHandwriting.forEach(function(hw) {
        list += '  ✍️ [' + hw.id.substring(0, 8) + '...] Size:' + hw.size + 'pt - "' + (hw.text.length > 30 ? hw.text.substring(0, 27) + '...' : hw.text) + '"\n';
        total++;
      });
    }
  });

  if (list === '') {
    list = 'No handwriting found.\n\nUse "Open Handwriting Panel" to add handwriting.';
  } else {
    list = 'Total: ' + total + ' handwriting elements\n' + list;
  }

  ui.alert('📋 All Handwriting', list, ui.ButtonSet.OK);
}


function clearAllHandwriting() {
  const ui = SlidesApp.getUi();

  const result = ui.alert(
    '⚠️ Clear All Handwriting',
    'This will remove all handwriting markers and associated data from Speaker Notes.\n\n' +
    'The visual text boxes will be cleared and made transparent.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) return;

  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();

  let count = 0;

  slides.forEach(function(slide) {
    const shapes = slide.getShapes();
    const notesShape = slide.getNotesPage().getSpeakerNotesShape();
    let notesModified = false;
    let notes = notesShape ? notesShape.getText().asString() : '';

    shapes.forEach(function(shape) {
       try {
          const description = shape.getDescription();
          if (description && description.startsWith(HANDWRITE_MARKER)) {
            // Clear shape
            shape.setDescription('');
            shape.getText().setText('');
            shape.getFill().setSolidFill('#FFFFFF', 0); // Transparent
            shape.getBorder().getLineFill().setSolidFill('#FFFFFF', 0); // No border

            // Mark notes for modification
            const handwriteId = description.substring(HANDWRITE_MARKER.length);
            const regex = new RegExp('\\[handwrite:' + handwriteId.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&') + ':[^\\]]+\\]\\n?', 'g');
            if (notes.match(regex)) {
                notes = notes.replace(regex, '');
                notesModified = true;
            }
            count++;
          }
        } catch (e) {
            Logger.log('Error clearing handwriting shape: ' + e);
        }
    });

    // Update notes if modified for the slide
    if (notesModified && notesShape) {
      notesShape.getText().setText(notes);
    }
  });

  ui.alert('✅ Cleared!', 'Removed data for ' + count + ' handwriting elements.', ui.ButtonSet.OK);
}


// ══════════════════════════════════════════════════════
// COORDINATE CONVERSION
// ══════════════════════════════════════════════════════

function emuToCm(emu, dimension) {
  let relativePosition;
  const EMU_PER_CM = 360000; // Standard conversion factor

  if (dimension === 'x') {
    // Direct conversion might be simpler and less prone to floating point issues
    // relativePosition = emu / SLIDE_WIDTH_EMU;
    // return relativePosition * BEAMER_WIDTH_CM;
     return emu / EMU_PER_CM;
  } else if (dimension === 'y') {
    // Convert EMU Y to CM Y (flipped)
    // relativePosition = emu / SLIDE_HEIGHT_EMU;
    // return (1 - relativePosition) * BEAMER_HEIGHT_CM;
    const yCmFromTop = emu / EMU_PER_CM;
    // Need Beamer height in CM, let's recalculate from standard EMU if needed, or use const
    // Assuming BEAMER_HEIGHT_CM is correctly defined relative to SLIDE_HEIGHT_EMU
    return BEAMER_HEIGHT_CM - yCmFromTop;
  } else if (dimension === 'w' || dimension === 'h') {
     // Convert width/height directly
     return emu / EMU_PER_CM;
  }
  return 0;
}

function convertRotation(slidesRotation) {
  // Beamer/TikZ rotation is typically anti-clockwise positive
  // Google Slides API rotation is clockwise positive
  return -slidesRotation;
}

// ══════════════════════════════════════════════════════
// AUTO-DETECT FROM COLORS
// ══════════════════════════════════════════════════════

function autoDetectFromColors() {
  const ui = SlidesApp.getUi();
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();

  const result = ui.alert(
    '🎨 Auto-Detect from Colors',
    'This will detect annotations from:\n\n' +
    '🔴 Red text = Circle annotation\n' +
    '🔵 Blue text = Box annotation\n' +
    '🟢 Green text = Underline annotation\n' +
    '🟡 Yellow background = Highlight\n\n' +
    'Plus read explanations from Speaker Notes:\n' +
    'Format: "Word: explanation"\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) return;

  let totalAnnotations = 0;
  let totalNotes = 0;

  slides.forEach(function(slide, index) {
    const coloredWords = detectColoredText(slide);
    const notesExplanations = parseNotesExplanations(slide);

    const merged = mergeAnnotationsWithNotes(coloredWords, notesExplanations);

    if (merged.length > 0) {
      saveMergedAnnotations(slide, merged); // Saves color annotations to speaker notes
      totalAnnotations += merged.length;
      totalNotes += merged.filter(function(a) { return a.note; }).length;
    }
  });

  ui.alert(
    '✅ Detection Complete!',
    'Found:\n' +
    '• ' + totalAnnotations + ' colored annotations\n' +
    '• ' + totalNotes + ' explanation notes\n\n' +
    'Annotations saved to Speaker Notes. Review and Export.',
    ui.ButtonSet.OK
  );
}

function detectColoredText(slide) {
  const colored = [];
  const shapes = slide.getShapes();

  shapes.forEach(function(shape) {
    try {
        const description = shape.getDescription();
        // Skip handwriting placeholders
        if (description && description.startsWith(HANDWRITE_MARKER)) {
          return;
        }

        const textRange = shape.getText();
        if (!textRange || textRange.isEmpty()) return; // Skip shapes without text

        const text = textRange.asString();
        if (text.trim().length === 0) return;

        const runs = textRange.getRuns();

        runs.forEach(function(run) {
          const runText = run.asString().trim();
          if (runText.length === 0) return;

          const style = run.getTextStyle();
          if (!style) return; // Skip if no style info

          // Check foreground color
          const foreColor = style.getForegroundColor();
          if (foreColor && foreColor.getColorType() === SlidesApp.ColorType.RGB) {
            const rgb = foreColor.asRgbColor();
            const r = rgb.getRed(); // Values are 0.0 to 1.0
            const g = rgb.getGreen();
            const b = rgb.getBlue();

            let type = null;
            // Use thresholds considering float values
            if (r > 0.8 && g < 0.4 && b < 0.4) { type = 'circle'; } // Red
            else if (r < 0.4 && g < 0.4 && b > 0.8) { type = 'box'; } // Blue
            else if (r < 0.4 && g > 0.8 && b < 0.4) { type = 'underline'; } // Green

            if (type) {
              colored.push({ text: runText, type: type });
            }
          }

          // Check background color (highlight)
          const bgColor = style.getBackgroundColor();
          if (bgColor && bgColor.getColorType() === SlidesApp.ColorType.RGB) {
            const rgb = bgColor.asRgbColor();
            const r = rgb.getRed();
            const g = rgb.getGreen();
            const b = rgb.getBlue();

            // Yellow highlight
            if (r > 0.8 && g > 0.8 && b < 0.4) {
              colored.push({ text: runText, type: 'underline' }); // Treat highlight as underline for now
            }
          }
        });
     } catch (e) {
         Logger.log('Error processing shape for color detection: ' + e);
     }
  });

  return colored;
}

function parseNotesExplanations(slide) {
  const notesShape = slide.getNotesPage().getSpeakerNotesShape();
  if (!notesShape) return {};

  const notes = notesShape.getText().asString();
  const explanations = {};
  const lines = notes.split('\n');

  lines.forEach(function(line) {
    line = line.trim();

    // Skip special lines
    if (line.startsWith('---') || line.startsWith('[annotation:') || line.startsWith('[handwrite:') || line.startsWith('💡')) {
       return;
    }
    if (line.length === 0) return;

    // Look for "Word: explanation" format
    const colonIndex = line.indexOf(':');
    if (colonIndex > 0 && colonIndex < line.length - 1) { // Ensure colon is not first or last char
      const word = line.substring(0, colonIndex).trim();
      const explanation = line.substring(colonIndex + 1).trim();

      if (word && explanation) {
        // Use first definition found for a word (lowercase match)
        if (!explanations[word.toLowerCase()]) {
            explanations[word.toLowerCase()] = explanation;
        }
      }
    }
  });

  return explanations;
}


function mergeAnnotationsWithNotes(coloredWords, explanations) {
  const merged = [];
  const addedAnnotations = new Set(); // Keep track of text+type already added to avoid duplicates from runs

  coloredWords.forEach(function(item) {
    const key = item.text + '|' + item.type;
    if (addedAnnotations.has(key)) return; // Skip duplicate runs with same color/highlight

    const word = item.text;
    const wordLower = word.toLowerCase();

    const annotation = {
      type: item.type,
      text: word,
      note: explanations[wordLower] || null // Add explanation if found
    };

    merged.push(annotation);
    addedAnnotations.add(key);
  });

  return merged;
}

// Function to save COLOR-detected annotations into speaker notes
function saveMergedAnnotations(slide, annotations) {
  const notesShape = slide.getNotesPage().getSpeakerNotesShape();
  if (!notesShape) {
      Logger.log('Cannot save annotations, speaker notes shape not found for slide.');
      return;
  }
  const existingNotes = notesShape.getText().asString();

  const lines = [];
  lines.push('--- AUTO-DETECTED COLOR ANNOTATIONS ---');

  let annotationIdCounter = 1; // Simple counter for IDs

  annotations.forEach(function(ann) {
    const id = annotationIdCounter.toString().padStart(2, '0');
    lines.push('[annotation:' + ann.type + ':' + id + ':' + ann.text + ']');
    if (ann.note) {
      lines.push('💡 Note [' + id + ']: ' + ann.note); // Associate note with ID
    }
    annotationIdCounter++;
  });

  lines.push('--- END AUTO-DETECTED COLOR ANNOTATIONS ---');
  lines.push(''); // Add a blank line separator

  // Remove old color annotation section, keep other notes (including handwriting and shape annotations)
  const cleanedNotes = existingNotes.replace(
    /--- AUTO-DETECTED COLOR ANNOTATIONS ---[\s\S]*?--- END AUTO-DETECTED COLOR ANNOTATIONS ---\n*/g,
    ''
  );

  const updatedNotes = newSection + '\n' + cleanedNotes.trim();

  notesShape.getText().setText(updatedNotes);
  Logger.log('Color annotations saved to speaker notes for slide.');
}


// ══════════════════════════════════════════════════════
// AUTO-DETECT FROM SHAPES
// ══════════════════════════════════════════════════════

function autoDetectFromShapes() {
  const ui = SlidesApp.getUi();
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();

  const result = ui.alert(
    '📐 Auto-Detect from Shapes',
    'This will detect annotations from shapes drawn around text:\n\n' +
    '⭕ Circles → Circle annotations\n' +
    '📦 Rectangles → Box annotations\n' +
    '📏 Long rectangles → Underline\n\n' +
    'Shapes must have:\n' +
    '• Transparent or very light fill\n' +
    '• Visible border (weight > 1pt)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) return;

  let processedSlides = 0;
  let totalDetected = 0;

  slides.forEach(function(slide, index) {
    try {
      const detectedCount = processSlideShapeAnnotations(slide);
      totalDetected += detectedCount;
      processedSlides++;
    } catch (error) {
      Logger.log('Error processing shape annotations on slide ' + (index + 1) + ': ' + error);
    }
  });

  ui.alert(
    '✅ Shape Detection Complete!',
    'Processed: ' + processedSlides + '/' + slides.length + ' slides\n' +
    'Detected: ' + totalDetected + ' shape annotations\n\n' +
    'Annotations saved to Speaker Notes. Review and Export.',
    ui.ButtonSet.OK
  );
}

// Renamed from processSlideAnnotations to be more specific
function processSlideShapeAnnotations(slide) {
  const shapes = slide.getShapes();
  const annotationShapes = [];
  const textElements = [];

  // 1. Classify shapes
  shapes.forEach(function(shape) {
     try {
        const description = shape.getDescription();
        // Skip handwriting placeholders
        if (description && description.startsWith(HANDWRITE_MARKER)) {
          return;
        }

        const textRange = shape.getText();

        if (textRange && !textRange.isEmpty()) {
          // It's a text element
          textElements.push({
            shape: shape,
            text: textRange.asString(),
            bounds: getElementBounds(shape),
            isList: isListElement(textRange.asString())
          });
        } else {
          // Potential annotation shape (no text)
          if (isAnnotationShape(shape)) {
            annotationShapes.push({
              shape: shape,
              type: detectAnnotationType(shape),
              bounds: getElementBounds(shape)
            });
          }
        }
      } catch (e) {
          Logger.log('Error classifying shape: ' + e);
      }
  });

  if (annotationShapes.length === 0 || textElements.length === 0) {
      Logger.log('No annotation shapes or text elements found on slide.');
      return 0; // Nothing to detect
  }


  // 2. Match annotation shapes with text elements
  const detectedAnnotations = [];
  let annotationIdCounter = 1; // Simple counter

  annotationShapes.forEach(function(annShape) {
    const matchedTexts = [];

    textElements.forEach(function(textEl) {
      // Calculate overlap between annotation shape and text element bounds
      const overlap = calculateOverlap(annShape.bounds, textEl.bounds);
      const textArea = textEl.bounds.width * textEl.bounds.height;
      const overlapRatio = textArea > 0 ? overlap / textArea : 0;

      // Require significant overlap (e.g., > 30% of text area)
      if (overlap > 0 && overlapRatio > 0.3) {
        matchedTexts.push({
          text: textEl.text,
          overlap: overlap,
          overlapRatio: overlapRatio, // Store ratio for better matching
          bounds: textEl.bounds
        });
      }
    });

    if (matchedTexts.length > 0) {
      // Sort matches by overlap ratio (highest first)
      matchedTexts.sort(function(a, b) { return b.overlapRatio - a.overlapRatio; });

      const bestMatch = matchedTexts[0];
      // Extract the relevant text (e.g., first word or line)
      const annotatedText = extractAnnotatedText(
        bestMatch.text,
        annShape.bounds,
        bestMatch.bounds
      );

      if (annotatedText) {
        const id = annotationIdCounter.toString().padStart(2, '0');
        detectedAnnotations.push({
          type: annShape.type,
          id: id,
          text: annotatedText
        });
        annotationIdCounter++;
      }
    }
  });

  // 3. Save detected shape annotations to Speaker Notes
  if (detectedAnnotations.length > 0) {
    const notesShape = slide.getNotesPage().getSpeakerNotesShape();
     if (!notesShape) {
        Logger.log('Cannot save shape annotations, speaker notes shape not found.');
        return detectedAnnotations.length; // Return count even if saving failed
     }
    const existingNotes = notesShape.getText().asString();

    const autoSection = buildAutoShapeAnnotationsSection(detectedAnnotations);
    const updatedNotes = mergeShapeNotes(existingNotes, autoSection);

    notesShape.getText().setText(updatedNotes);
    Logger.log(detectedAnnotations.length + ' shape annotations saved to speaker notes.');
  }

  return detectedAnnotations.length; // Return number of annotations detected
}


function isAnnotationShape(shape) {
  try {
      const shapeType = shape.getShapeType();

      // Must be a supported geometric shape
      if (shapeType !== SlidesApp.ShapeType.ELLIPSE &&
          shapeType !== SlidesApp.ShapeType.RECTANGLE &&
          shapeType !== SlidesApp.ShapeType.ROUND_RECTANGLE) {
        return false;
      }

      // Must not contain text (handled separately)
      const textRange = shape.getText();
      if (textRange && !textRange.isEmpty()) {
          return false;
      }

      // Check fill: must be NONE or very transparent
      const fill = shape.getFill();
      const fillType = fill.getFillType();
      if (fillType === SlidesApp.FillType.SOLID) {
        const solidFill = fill.getSolidFill();
        // Check alpha (transparency). Alpha 1.0 is opaque, 0.0 is fully transparent.
        if (solidFill.getAlpha() === null || solidFill.getAlpha() > 0.1) { // Allow slight transparency
            return false;
        }
      } else if (fillType !== SlidesApp.FillType.NONE) {
          return false; // Only allow solid transparent or no fill
      }

      // Check border: must have a visible border
      const border = shape.getBorder();
      if (border.getLineFill().getFillType() === SlidesApp.FillType.NONE || border.getWeight() < 0.5) { // Require a minimal border
          return false;
      }

      // If all checks pass, it's likely an annotation shape
      return true;
   } catch (e) {
       Logger.log('Error in isAnnotationShape: ' + e);
       return false; // Treat errors as non-annotation shapes
   }
}

function detectAnnotationType(shape) {
  const shapeType = shape.getShapeType();

  if (shapeType === SlidesApp.ShapeType.ELLIPSE) {
    return 'circle';
  }

  if (shapeType === SlidesApp.ShapeType.RECTANGLE || shapeType === SlidesApp.ShapeType.ROUND_RECTANGLE) {
    const width = shape.getWidth();
    const height = shape.getHeight();

    // Check aspect ratio for underline (long and thin)
    if (width > 0 && height > 0 && width / height > 5 && height < 300000) { // EMU threshold for height
      return 'underline';
    }
    return 'box';
  }

  return 'box'; // Default
}

function getElementBounds(element) {
  try {
      const transform = element.getTransform();
      const size = element.getSize(); // Use getSize() for consistency

      const left = transform.getTranslateX() || 0;
      const top = transform.getTranslateY() || 0;
      const width = size.getWidth() || 0;
      const height = size.getHeight() || 0;

      return {
        left: left,
        top: top,
        right: left + width,
        bottom: top + height,
        width: width,
        height: height,
        centerX: left + width / 2,
        centerY: top + height / 2
      };
   } catch (e) {
       Logger.log('Error getting element bounds: ' + e);
       return { left: 0, top: 0, right: 0, bottom: 0, width: 0, height: 0, centerX: 0, centerY: 0 };
   }
}


function calculateOverlap(bounds1, bounds2) {
  // Find the overlapping rectangle coordinates
  const overlapLeft = Math.max(bounds1.left, bounds2.left);
  const overlapRight = Math.min(bounds1.right, bounds2.right);
  const overlapTop = Math.max(bounds1.top, bounds2.top);
  const overlapBottom = Math.min(bounds1.bottom, bounds2.bottom);

  // Check if there is an actual overlap
  if (overlapLeft < overlapRight && overlapTop < overlapBottom) {
    const overlapWidth = overlapRight - overlapLeft;
    const overlapHeight = overlapBottom - overlapTop;
    return overlapWidth * overlapHeight; // Return area of overlap
  }

  return 0; // No overlap
}

// Improved heuristic to extract annotated text
function extractAnnotatedText(fullText, annBounds, textBounds) {
  if (!fullText || fullText.trim().length === 0) return null;

  const lines = fullText.split('\n');

  // Heuristic 1: If annotation is small, likely a word or short phrase
  if (annBounds.width < 300000 && annBounds.height < 300000) { // Small EMU threshold
     const words = fullText.match(/\b\w+\b/g); // Extract words
     if (words) {
         // Maybe find word closest to center of annotation? (More complex)
         // For now, return the first meaningful word
         const firstWord = words.find(w => w.length > 2);
         if (firstWord) return firstWord;
     }
  }

  // Heuristic 2: If annotation covers most of the text box width, return first line
  if (annBounds.width > textBounds.width * 0.7 && lines.length > 0) {
      const firstLine = lines[0].trim();
      if (firstLine.length > 0 && firstLine.length < 100) { // Limit line length
          return firstLine;
      }
  }

  // Fallback: return the first meaningful part (e.g., first 50 chars)
  const cleaned = fullText.trim().substring(0, 50);
  return cleaned.length > 0 ? cleaned : null;
}

// Function to build the section for SHAPE-detected annotations
function buildAutoShapeAnnotationsSection(annotations) {
  if (!annotations || annotations.length === 0) return '';

  const lines = [];
  lines.push('--- AUTO-DETECTED SHAPE ANNOTATIONS ---');

  annotations.forEach(function(ann) {
    lines.push('[annotation:' + ann.type + ':' + ann.id + ':' + ann.text + ']');
  });

  lines.push('--- END AUTO-DETECTED SHAPE ANNOTATIONS ---');
  lines.push(''); // Blank line separator
  lines.push('💡 Review shape annotations above.');
  lines.push('');

  return lines.join('\n');
}

// Function to merge SHAPE-detected notes, keeping other sections
function mergeShapeNotes(existingNotes, autoSection) {
  // Remove only the old SHAPE annotation section
  const cleanedNotes = existingNotes.replace(
    /--- AUTO-DETECTED SHAPE ANNOTATIONS ---[\s\S]*?--- END AUTO-DETECTED SHAPE ANNOTATIONS ---\n*/g,
    ''
  );

  // Add the new section at the top (or bottom, depending on preference)
  return autoSection + '\n' + cleanedNotes.trim();
}


// ══════════════════════════════════════════════════════
// MANUAL ANNOTATION SIDEBAR
// ══════════════════════════════════════════════════════

function showAnnotationSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('AnnotationSidebar')
    .setTitle('🖊️ Quick Annotate')
    .setWidth(320);

  SlidesApp.getUi().showSidebar(html);
}

// Function called from AnnotationSidebar.html to add manual annotation to notes
function addManualAnnotation(text, type) {
  const presentation = SlidesApp.getActivePresentation();
  let slide;

  // Get current slide or fallback to first
  const selection = presentation.getSelection();
  if (selection && selection.getSelectionType() === SlidesApp.SelectionType.PAGE) {
    slide = selection.getCurrentPage().asSlide();
  } else {
    const slides = presentation.getSlides();
    slide = slides.length > 0 ? slides[0] : null;
  }

  if (!slide) throw new Error("Could not determine the current slide.");

  const notesShape = slide.getNotesPage().getSpeakerNotesShape();
   if (!notesShape) throw new Error("Could not access speaker notes for this slide.");

  const existingNotes = notesShape.getText().asString();

  // Find the next available ID number for standard annotations
  const existingAnnotations = parseAnnotations(existingNotes);
  let nextIdNum = 0;
  existingAnnotations.forEach(ann => {
      const num = parseInt(ann.id);
      if (!isNaN(num) && num > nextIdNum) {
          nextIdNum = num;
      }
  });
  const idStr = (nextIdNum + 1).toString().padStart(2, '0'); // Simple sequential ID

  const newAnnotation = '[annotation:' + type + ':' + idStr + ':' + text + ']';

  const updatedNotes = (existingNotes.trim().length === 0) ? newAnnotation : existingNotes + '\n' + newAnnotation;

  notesShape.getText().setText(updatedNotes);

  return {
    success: true,
    id: idStr,
    type: type,
    text: text
  };
}

// Function called from AnnotationSidebar.html to add manual handwriting note
// Note: This adds a note but doesn't link it to a visual element.
// Use the Handwriting Panel for visual placement.
function addManualHandwriting(text) {
  const presentation = SlidesApp.getActivePresentation();
   let slide;

  // Get current slide or fallback to first
  const selection = presentation.getSelection();
  if (selection && selection.getSelectionType() === SlidesApp.SelectionType.PAGE) {
    slide = selection.getCurrentPage().asSlide();
  } else {
    const slides = presentation.getSlides();
    slide = slides.length > 0 ? slides[0] : null;
  }

  if (!slide) throw new Error("Could not determine the current slide.");

  const notesShape = slide.getNotesPage().getSpeakerNotesShape();
   if (!notesShape) throw new Error("Could not access speaker notes for this slide.");

  const existingNotes = notesShape.getText().asString();

  // Generate a unique ID for manual notes
  const timestamp = new Date().getTime();
  const idStr = 'note-' + timestamp + '-' + Math.random().toString(36).substring(2, 5);
  const defaultSize = 10; // Default size for notes added this way

  const newNote = '[handwrite:' + idStr + ':' + defaultSize + ':' + text + ']';

  const updatedNotes = (existingNotes.trim().length === 0) ? newNote : existingNotes + '\n' + newNote;

  notesShape.getText().setText(updatedNotes);

  return {
    success: true,
    id: idStr,
    text: text
  };
}

// Gets *all* annotations (color, shape, manual) and *manual* handwriting notes from current slide notes
function getCurrentSlideAnnotations() {
  const presentation = SlidesApp.getActivePresentation();
  let slide;

  // Get current slide or fallback to first
  const selection = presentation.getSelection();
  if (selection && selection.getSelectionType() === SlidesApp.SelectionType.PAGE) {
    slide = selection.getCurrentPage().asSlide();
  } else {
    const slides = presentation.getSlides();
    slide = slides.length > 0 ? slides[0] : null;
  }

  if (!slide) return { annotations: [], handwriting: [] }; // No slide context

  const notesShape = slide.getNotesPage().getSpeakerNotesShape();
  const notes = notesShape ? notesShape.getText().asString() : '';

  return {
    annotations: parseAnnotations(notes), // Standard annotations
    handwriting: parseHandwritingNotes(notes) // Manual/unlinked handwriting notes
  };
}

// Deletes a standard annotation line from speaker notes
function deleteAnnotation(id, type, text) {
   const presentation = SlidesApp.getActivePresentation();
   let slide;
   // ... (get current slide or fallback) ...
   if (!slide) throw new Error("No slide context for deletion.");

   const notesShape = slide.getNotesPage().getSpeakerNotesShape();
   if (!notesShape) throw new Error("Cannot access speaker notes.");

   const existingNotes = notesShape.getText().asString();

   // Need robust way to find the exact line, escaping text is crucial
   const escapedText = text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // Escape regex special chars
   // Match the specific annotation line, including potential surrounding whitespace/newlines
   const pattern = new RegExp('^\\s*\\[annotation:' + type + ':' + id + ':' + escapedText + '\\]\\s*\\n?', 'gm');

   const updatedNotes = existingNotes.replace(pattern, '').trim();

   notesShape.getText().setText(updatedNotes);
   Logger.log('Deleted annotation: [' + type + ':' + id + ':' + text + ']');

   return { success: true };
}

// Deletes a manual/unlinked handwriting note line from speaker notes
function deleteHandwritingNote(id) {
    const presentation = SlidesApp.getActivePresentation();
    let slide;
    // ... (get current slide or fallback) ...
    if (!slide) throw new Error("No slide context for deletion.");

    const notesShape = slide.getNotesPage().getSpeakerNotesShape();
    if (!notesShape) throw new Error("Cannot access speaker notes.");

    const existingNotes = notesShape.getText().asString();

    // Match the specific handwriting line using its unique ID
    const regex = new RegExp('^\\s*\\[handwrite:' + id.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&') + ':[^\\]]+\\]\\s*\\n?', 'gm');

    const updatedNotes = existingNotes.replace(regex, '').trim();

    notesShape.getText().setText(updatedNotes);
    Logger.log('Deleted handwriting note: [' + id + ']');

    return { success: true };
}

// ══════════════════════════════════════════════════════
// PARSING HELPERS
// ══════════════════════════════════════════════════════

// Parses standard annotations: [annotation:TYPE:ID:TEXT]
function parseAnnotations(notes) {
  if (!notes) return [];
  const annotations = [];
  // Regex: Find type (word), ID (digits), text (anything until ']')
  const regex = /\[annotation:(circle|cross|underline|box):(\d+):([^\]]+)\]/g;
  let match;

  while ((match = regex.exec(notes)) !== null) {
    annotations.push({
      type: match[1],
      id: match[2],
      text: match[3].trim()
    });
  }
  return annotations;
}

// Parses handwriting notes: [handwrite:ID:SIZE:TEXT]
function parseHandwritingNotes(notes) {
  if (!notes) return [];
  const handwriting = [];
   // Regex: Find ID (starts with 'hw-' or 'note-'), size (digits), text (anything until ']')
  const regex = /\[handwrite:(hw-[\w-]+|note-[\w-]+):(\d+):([^\]]+)\]/g;
  let match;

  while ((match = regex.exec(notes)) !== null) {
    handwriting.push({
      id: match[1], // Includes 'hw-' or 'note-' prefix
      size: match[2],
      text: match[3].trim()
    });
  }
  return handwriting;
}

function isListElement(text) {
  if (!text) return false;
  // Check for common bullet point starts or numbered list starts
  return text.includes('\n•') ||
         text.includes('\n-') ||
         text.includes('\n*') ||
         /^\s*[•\-\*]/.test(text) || // Bullet at the very beginning
         /^\s*\d+\.\s+/.test(text);   // Numbered list item like "1. "
}


// ══════════════════════════════════════════════════════
// PREVIEW & CLEAR
// ══════════════════════════════════════════════════════

function previewCapturedData() {
  const ui = SlidesApp.getUi();
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();

  let preview = '';
  let totalAnnotations = 0;
  let totalHandwriting = 0; // Counts visually placed handwriting

  slides.forEach(function(slide, index) {
    let slideContent = '';
    const notes = slide.getNotesPage().getSpeakerNotesShape()?.getText()?.asString() || '';
    const annotations = parseAnnotations(notes); // Get standard annotations from notes
    const shapes = slide.getShapes();
    let slideHasHandwriting = false;

    // Add standard annotations to preview
    annotations.forEach(function(ann) {
      slideContent += '  🔵 [' + ann.type + ':' + ann.id + '] "' + ann.text + '"\n';
      totalAnnotations++;
    });

    // Check for visually placed handwriting
    shapes.forEach(function(shape) {
       try {
          const description = shape.getDescription();
          if (description && description.startsWith(HANDWRITE_MARKER)) {
            const handwriteId = description.substring(HANDWRITE_MARKER.length);
            const regex = new RegExp('\\[handwrite:' + handwriteId.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&') + ':(\\d+):([^\\]]+)\\]');
            const match = notes.match(regex);
            const text = match ? match[2] : shape.getText().asString(); // Fallback to preview text
            slideContent += '  ✏️ [Handwriting:' + handwriteId.substring(0,6) + '] "' + text + '"\n';
            totalHandwriting++;
            slideHasHandwriting = true;
          }
        } catch(e){}
    });

    // Add section header if there's content for the slide
    if (slideContent) {
        preview += '\n📄 Slide ' + (index + 1) + ':\n' + slideContent;
    }
  });

  if (!preview) {
    preview = 'No annotations or handwriting captured yet.\n\nUse annotation tools or the Handwriting Panel.';
  } else {
    preview = '📊 Summary:\n' +
              '  Total Annotations: ' + totalAnnotations + '\n' +
              '  Total Handwriting: ' + totalHandwriting + '\n' + preview;
  }

  ui.alert('📋 Captured Data Preview', preview, ui.ButtonSet.OK);
}


function clearCapturedData() {
  const ui = SlidesApp.getUi();
  const result = ui.alert(
    '⚠️ Clear Auto-Detected Annotations Only',
    'This will remove annotation sections created by:\n' +
    '  - Auto-Detect from Colors\n' +
    '  - Auto-Detect from Shapes\n\n' +
    'Manually added annotations and ALL handwriting data will be preserved.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) return;

  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  let clearedCount = 0;

  slides.forEach(function(slide) {
    const notesShape = slide.getNotesPage().getSpeakerNotesShape();
    if (!notesShape) return;

    let existingNotes = notesShape.getText().asString();
    let notesChanged = false;

    // Remove color annotations section
    const colorRegex = /--- AUTO-DETECTED COLOR ANNOTATIONS ---[\s\S]*?--- END AUTO-DETECTED COLOR ANNOTATIONS ---\n*/g;
    if (existingNotes.match(colorRegex)) {
       existingNotes = existingNotes.replace(colorRegex, '');
       notesChanged = true;
       clearedCount++;
    }

    // Remove shape annotations section
    const shapeRegex = /--- AUTO-DETECTED SHAPE ANNOTATIONS ---[\s\S]*?--- END AUTO-DETECTED SHAPE ANNOTATIONS ---\n*/g;
     if (existingNotes.match(shapeRegex)) {
        existingNotes = existingNotes.replace(shapeRegex, '');
        notesChanged = true;
        if (!notesChanged) clearedCount++; // Count only once per slide if both existed
     }

    // Update notes only if something was removed
    if (notesChanged) {
      notesShape.getText().setText(existingNotes.trim());
    }
  });

  ui.alert('✅ Cleared!', 'Removed auto-detected annotation sections from ' + clearedCount + ' slides.', ui.ButtonSet.OK);
}


// ══════════════════════════════════════════════════════
// EXPORT TO LATEX
// ══════════════════════════════════════════════════════

function exportToNagwaLatex() {
  const ui = SlidesApp.getUi();
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();

  if (slides.length === 0) {
    ui.alert('❌ Error', 'No slides found in presentation!', ui.ButtonSet.OK);
    return;
  }

  // Start processing notification
  ui.showSidebar(HtmlService.createHtmlOutput('<p>Processing ' + slides.length + ' slides... Please wait.</p>').setTitle('Exporting...'));


  try {
    const config = getSessionConfig();
    const latex = generateNagwaLatex(presentation, slides, config);

    const props = PropertiesService.getDocumentProperties();
    props.setProperty('latest_latex', latex);
    props.setProperty('latest_latex_timestamp', new Date().toString());

    // Close processing sidebar before showing results
     ui.showSidebar(null); // This might not work as expected, modal dialog is better

    showLatexPreview(latex);

  } catch (error) {
     ui.showSidebar(null); // Close processing on error too
    Logger.log('Export error: ' + error + '\n' + error.stack);
    ui.alert(
      '❌ Export Failed',
      'Error: ' + error.message + '\n\n' +
      'Check View → Logs for more details.',
      ui.ButtonSet.OK
    );
  }
}

function generateNagwaLatex(presentation, slides, config) {
  const latex = [];

  // Preamble
  latex.push(
    '\\documentclass[' + config.documentClass + ', ' +
    'nagwalang=' + config.language + ', ' +
    'numerals=' + config.numerals + ', ' +
    'directions=' + config.directions + ']{nagwa}'
  );
  latex.push('');
  latex.push('\\usepackage{roughnotation}');
  latex.push('\\usepackage{tikz}'); // Needed for absolute positioning potentially
  latex.push('\\usepackage{graphicx}');
  latex.push('\\usepackage{amsmath}');
  latex.push('\\usepackage{amssymb}');
  latex.push('');
  latex.push('\\begin{document}');
  latex.push('');

  // Metadata
  latex.push('    \\metasessionID{' + config.sessionID + '}');
  latex.push('    \\sessioncountry{' + config.country + '}');
  latex.push('    \\subject{' + config.subject + '}');
  latex.push('    \\languageofinstruction{' + config.language + '}');
  latex.push('    \\grade{' + config.grade + '}');
  latex.push('    \\term{' + config.term + '}');
  latex.push('    \\sessiontitle{' + escapeLatex(config.sessionTitle) + '}'); // Escape title
  latex.push('');

  // Slides
  slides.forEach(function(slide, index) {
    try {
        const slideLatex = convertSlideToLatex(slide, index);
        latex.push(...slideLatex); // Add all lines from slide conversion
        latex.push(''); // Add blank line between slides
    } catch (e) {
        Logger.log('Error converting slide ' + (index + 1) + ': ' + e);
        latex.push('% --- ERROR CONVERTING SLIDE ' + (index + 1) + ': ' + e.message + ' ---');
        latex.push('');
    }
  });

  latex.push('\\end{document}');

  return latex.join('\n');
}


function convertSlideToLatex(slide, slideIndex) {
  const latex = [];
  const slideId = slide.getObjectId();
  Logger.log('Processing Slide ' + (slideIndex + 1) + ' (ID: ' + slideId + ')');

  const title = extractSlideTitle(slide);
  const shapes = slide.getShapes();
  const images = slide.getImages(); // Get images separately
  const tables = slide.getTables(); // Get tables separately

  const notes = slide.getNotesPage().getSpeakerNotesShape()?.getText()?.asString() || '';
  const annotationsInNotes = parseAnnotations(notes); // Get all standard annotations

  const handwritingPlaceholders = [];
  const normalTextElements = [];
  const processedElementIds = new Set(); // To avoid processing elements twice

  // --- Step 1: Identify Handwriting Placeholders ---
  shapes.forEach(function(shape) {
    try {
      const description = shape.getDescription();
      if (description && description.startsWith(HANDWRITE_MARKER)) {
        const handwriteId = description.substring(HANDWRITE_MARKER.length);
        const hwData = extractHandwritingData(shape, handwriteId, notes);
        if (hwData) {
          handwritingPlaceholders.push(hwData);
          processedElementIds.add(shape.getObjectId()); // Mark as processed
          Logger.log('Identified Handwriting: ' + handwriteId);
        } else {
           Logger.log('Handwriting data missing in notes for ID: ' + handwriteId);
        }
      }
     } catch (e) { Logger.log('Error checking shape for handwriting: ' + e); }
  });

  // --- Step 2: Identify Normal Text Elements ---
   shapes.forEach(function(shape) {
      if (processedElementIds.has(shape.getObjectId())) return; // Skip if already processed as handwriting

      try {
         const textRange = shape.getText();
         if (textRange && !textRange.isEmpty()) {
             normalTextElements.push({
               text: textRange.asString(), // Get full text content
               isList: isListElement(textRange.asString())
             });
             processedElementIds.add(shape.getObjectId()); // Mark as processed
         }
       } catch (e) { Logger.log('Error processing shape for text: ' + e); }
   });


  // --- Step 3: Generate LaTeX Output ---
  latex.push('    % --- Slide ' + (slideIndex + 1) + ' / ID: ' + slideId + ' --- ');
  latex.push('    \\slidestandalone{' + escapeLatex(title) + '}{%');

  // Add Normal Text Content (with annotations applied)
  if (normalTextElements.length > 0) {
    const contentLatex = generateSlideContent(normalTextElements, annotationsInNotes);
    latex.push(...contentLatex.map(line => '        ' + line)); // Indent content
  } else {
      latex.push('        % No standard text content detected.');
  }

  // Add Images (Placeholder - needs implementation for positioning)
  if (images.length > 0) {
      latex.push('');
      latex.push('        % --- Images (Add positioning logic) ---');
      images.forEach((img, idx) => {
          // Placeholder: just include image, needs coordinates
          // const imgBounds = getElementBounds(img);
          // const imgX = emuToCm(imgBounds.left, 'x').toFixed(2);
          // const imgY = emuToCm(imgBounds.top, 'y').toFixed(2);
          // const imgW = emuToCm(imgBounds.width, 'w').toFixed(2);
          latex.push('        \\includegraphics[width=0.5\\textwidth]{image_' + slideIndex + '_' + idx + '.png} % TODO: Add position');
      });
  }

   // Add Tables (Placeholder - needs implementation)
   if (tables.length > 0) {
       latex.push('');
       latex.push('        % --- Tables (Add conversion logic) ---');
       tables.forEach((table, idx) => {
           latex.push('        % Table ' + idx + ' data goes here');
       });
   }


  // Add Handwriting Elements
  if (handwritingPlaceholders.length > 0) {
    latex.push('');
    latex.push('        % --- Handwriting Elements ---');
    handwritingPlaceholders.forEach(function(hw, index) {
      // Use the unique ID from the note, not a simple counter
      const hwCommand = '        \\handwrite[' + hw.id + ',' +
                        hw.x + ',' + hw.y + ',' +
                        hw.rotation + ',' + hw.size + ']{' +
                        escapeLatex(hw.text) + '}'; // Ensure text is escaped
      latex.push(hwCommand);
    });
  }

  latex.push('    }'); // End \slidestandalone content

  return latex;
}


function extractHandwritingData(shape, handwriteId, notes) {
  try {
      const transform = shape.getTransform();
      if (!transform) return null; // Cannot get position if no transform

      const xEmu = transform.getTranslateX() || 0; // Default to 0 if null
      const yEmu = transform.getTranslateY() || 0;

      // Convert position
      const xCm = emuToCm(xEmu, 'x');
      const yCm = emuToCm(yEmu, 'y');

      // Convert rotation
      const rotationDegrees = transform.getRotation() || 0;
      const latexRotation = convertRotation(rotationDegrees);

      // Find the corresponding data in notes
      const regex = new RegExp('\\[handwrite:' + handwriteId.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&') + ':(\\d+):([^\\]]+)\\]');
      const match = notes.match(regex);

      if (!match) {
        Logger.log('Warning: Handwriting data not found in notes for ID: ' + handwriteId);
        return null; // Data mismatch or missing
      }

      const size = match[1];
      const text = match[2];

      return {
        id: handwriteId, // Use the unique ID
        x: xCm.toFixed(2),
        y: yCm.toFixed(2),
        rotation: latexRotation.toFixed(1),
        size: size,
        text: text
      };
   } catch (e) {
       Logger.log('Error extracting handwriting data for ID ' + handwriteId + ': ' + e);
       return null;
   }
}

// Generates LaTeX for normal text elements, applying annotations from notes
function generateSlideContent(textElements, annotationsFromNotes) {
  const content = [];
  let appliedAnnotationIds = new Set(); // Track which annotations were applied

  // Determine if the slide content should be treated as a list
  const hasLists = textElements.some(el => el.isList);

  if (hasLists) {
    content.push('\\begin{itemize}');
    textElements.forEach(element => {
      let currentItemsText = element.text;

      // Apply annotations *before* splitting into items if possible
      annotationsFromNotes.forEach(ann => {
          if (!appliedAnnotationIds.has(ann.id) && currentItemsText.includes(ann.text)) {
              currentItemsText = currentItemsText.replace(ann.text, formatAnnotation(ann));
              appliedAnnotationIds.add(ann.id); // Mark as applied
          }
      });

      if (element.isList) {
        const items = parseListItems(currentItemsText); // Parse potentially modified text
        items.forEach(item => {
          // Re-check annotations in case replace didn't catch fragmented text within items
          let finalItemText = item;
           annotationsFromNotes.forEach(ann => {
              if (!appliedAnnotationIds.has(ann.id) && finalItemText.includes(ann.text)) {
                  finalItemText = finalItemText.replace(ann.text, formatAnnotation(ann));
                  appliedAnnotationIds.add(ann.id);
              }
           });
          content.push('    \\item ' + escapeLatex(finalItemText)); // Escape the final item text
        });
      } else {
        // Treat non-list text block as a single item if mixed content
        if (currentItemsText.trim().length > 0) {
          content.push('    \\item ' + escapeLatex(currentItemsText));
        }
      }
    });
    content.push('\\end{itemize}');
  } else {
    // Treat as regular paragraphs
    textElements.forEach(element => {
      if (element.text.trim().length > 0) {
          let paragraphText = element.text;
           // Apply annotations
          annotationsFromNotes.forEach(ann => {
              if (!appliedAnnotationIds.has(ann.id) && paragraphText.includes(ann.text)) {
                  paragraphText = paragraphText.replace(ann.text, formatAnnotation(ann));
                  appliedAnnotationIds.add(ann.id);
              }
          });
        content.push(escapeLatex(paragraphText)); // Escape the final paragraph text
        content.push(''); // Add blank line between paragraphs
      }
    });
  }

  // Log any annotations that were in notes but not applied (optional debug)
  annotationsFromNotes.forEach(ann => {
      if (!appliedAnnotationIds.has(ann.id)) {
          Logger.log('Warning: Annotation [' + ann.id + '] "' + ann.text + '" from notes was not applied to slide content.');
      }
  });


  return content;
}

// Helper to format an annotation object into LaTeX command string
function formatAnnotation(ann) {
    // Note: Text inside annotation should NOT be escaped by escapeLatex here,
    // as LaTeX commands handle internal text differently. Let LaTeX handle it.
    // However, the ID and type are safe.
    const escapedText = ann.text; // Use raw text for inside the command

    if (ann.type === 'circle') {
      return '\\annotateCircle[' + ann.id + ']{' + escapedText + '}';
    } else if (ann.type === 'cross') {
      return '\\annotateCross[' + ann.id + ']{' + escapedText + '}';
    } else if (ann.type === 'underline') {
      return '\\annotateUnderline[' + ann.id + ']{' + escapedText + '}';
    } else if (ann.type === 'box') {
      return '\\annotateBox[' + ann.id + ']{' + escapedText + '}';
    } else {
      return escapedText; // Fallback if type is unknown
    }
}


function parseListItems(text) {
  if (!text) return [];
  const lines = text.split('\n');
  const items = [];

  lines.forEach(function(line) {
    // Trim leading/trailing whitespace
    let itemText = line.trim();
    if (itemText.length === 0) return;

    // Remove common list markers (bullet points or numbers)
    // More robust regex: matches optional whitespace, marker, then whitespace
    itemText = itemText.replace(/^\s*[•\-\*]\s*/, '');
    itemText = itemText.replace(/^\s*\d+\.\s+/, '');

    // Add only if there's remaining text
    if (itemText.length > 0) {
      items.push(itemText);
    }
  });

  return items;
}

// This function is deprecated by the new export logic which reads annotations from notes.
// function applyAnnotationsToText(text, annotations) { ... }


function extractSlideTitle(slide) {
  const shapes = slide.getShapes();

  // Priority 1: Check for official TITLE placeholder
  for (let i = 0; i < shapes.length; i++) {
     try {
        const shape = shapes[i];
        if (shape.getPlaceholderType && // Check if function exists
            (shape.getPlaceholderType() === SlidesApp.PlaceholderType.TITLE ||
             shape.getPlaceholderType() === SlidesApp.PlaceholderType.CENTERED_TITLE)) {
            const titleText = shape.getText()?.asString()?.trim();
            if (titleText) return titleText;
        }
      } catch (e) { Logger.log('Error checking placeholder: ' + e); }
  }

  // Priority 2: Find the first non-placeholder, non-handwriting text box with content
  for (let i = 0; i < shapes.length; i++) {
     try {
        const shape = shapes[i];
        const description = shape.getDescription();
        // Skip handwriting placeholders
        if (description && description.startsWith(HANDWRITE_MARKER)) continue;
        // Skip placeholders unless it's the title (handled above)
        if (shape.getPlaceholderType && shape.getPlaceholderType() !== SlidesApp.PlaceholderType.NONE) continue;

        const text = shape.getText()?.asString()?.trim();
        // Return first reasonably short text found
        if (text && text.length > 0 && text.length < 150) {
          return text;
        }
      } catch (e) { Logger.log('Error checking shape for title: ' + e); }
  }

  // Fallback: Generic title
  return 'Untitled Slide';
}


function escapeLatex(text) {
    if (text === null || text === undefined) {
        return '';
    }
    let result = String(text); // Ensure it's a string

    // Order matters: escape backslash first
    result = result.replace(/\\/g, '\\textbackslash{}');

    // Escape other special characters
    result = result.replace(/&/g, '\\&');
    result = result.replace(/%/g, '\\%');
    result = result.replace(/\$/g, '\\$');
    result = result.replace(/#/g, '\\#');
    result = result.replace(/_/g, '\\_');
    result = result.replace(/\{/g, '\\{');
    result = result.replace(/\}/g, '\\}');
    result = result.replace(/~/g, '\\textasciitilde{}');
    result = result.replace(/\^/g, '\\textasciicircum{}');
    // Add escaping for < and > if needed, depending on context
    // result = result.replace(/</g, '\\textless{}');
    // result = result.replace(/>/g, '\\textgreater{}');

    return result;
}


function showLatexPreview(latex) {
  const lines = latex.split('\n').length;
  const chars = latex.length;
  const slideMatches = latex.match(/\\slidestandalone/g);
  const slideCount = slideMatches ? slideMatches.length : 0;
  const hwMatches = latex.match(/\\handwrite\[/g);
  const hwCount = hwMatches ? hwMatches.length : 0;

  // Improved HTML for better UI and clarity
  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Roboto', Arial, sans-serif; padding: 20px; background-color: #f8f9fa; margin: 0; display: flex; flex-direction: column; height: 100vh; }
        .header { background: linear-gradient(135deg, #673ab7 0%, #3f51b5 100%); color: white; padding: 20px; border-radius: 8px 8px 0 0; margin: -20px -20px 20px -20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h2 { margin: 0; font-size: 24px; font-weight: 400; }
        .subtitle { opacity: 0.9; margin-top: 4px; font-size: 14px; }
        .stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(100px, 1fr)); gap: 12px; margin-bottom: 20px; background-color: #fff; padding: 16px; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
        .stat { text-align: center; }
        .stat-value { font-size: 24px; font-weight: 500; color: #3f51b5; }
        .stat-label { font-size: 11px; color: #5f6368; margin-top: 4px; text-transform: uppercase; letter-spacing: 0.5px; }
        .main-content { flex-grow: 1; display: flex; flex-direction: column; min-height: 0; }
        textarea { width: 100%; flex-grow: 1; font-family: 'Courier New', monospace; font-size: 12px; padding: 12px; border: 1px solid #ddd; border-radius: 8px; box-sizing: border-box; resize: none; margin-bottom: 16px; background-color: #fff; }
        .buttons { display: flex; gap: 10px; flex-wrap: wrap; }
        button { padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 500; transition: all 0.2s ease; }
        .btn-primary { background-color: #3f51b5; color: white; }
        .btn-primary:hover { background-color: #303f9f; transform: translateY(-1px); box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        .btn-secondary { background-color: #4caf50; color: white; }
        .btn-secondary:hover { background-color: #388e3c; transform: translateY(-1px); box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        .btn-tertiary { background-color: #757575; color: white; margin-left: auto; } /* Push close button to the right */
        .btn-tertiary:hover { background-color: #616161; }
        .notification { position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%); background-color: #333; color: white; padding: 10px 20px; border-radius: 4px; font-size: 13px; box-shadow: 0 2px 5px rgba(0,0,0,0.2); z-index: 1000; opacity: 0; transition: opacity 0.3s ease; }
        .notification.show { opacity: 1; }
      </style>
    </head>
    <body>
      <div class="header">
        <h2>✅ LaTeX Export Complete!</h2>
        <div class="subtitle">Ready to compile with Nagwa LaTeX</div>
      </div>

      <div class="stats">
        <div class="stat"><div class="stat-value">${slideCount}</div><div class="stat-label">Slides</div></div>
        <div class="stat"><div class="stat-value">${hwCount}</div><div class="stat-label">Handwriting</div></div>
        <div class="stat"><div class="stat-value">${lines.toLocaleString()}</div><div class="stat-label">Lines</div></div>
        <div class="stat"><div class="stat-value">${chars.toLocaleString()}</div><div class="stat-label">Characters</div></div>
      </div>

      <div class="main-content">
        <textarea id="code" readonly>${latex.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</textarea>

        <div class="buttons">
          <button class="btn-primary" onclick="copyCode()">📋 Copy to Clipboard</button>
          <button class="btn-secondary" onclick="downloadFile()">💾 Download .tex File</button>
          <button class="btn-tertiary" onclick="google.script.host.close()">Close</button>
        </div>
      </div>

      <div id="notification" class="notification"></div>

      <script>
        function copyCode() {
          const textarea = document.getElementById('code');
          textarea.select();
          textarea.setSelectionRange(0, 99999); // For mobile devices
          try {
            document.execCommand('copy');
            showNotification('✅ Copied to clipboard!');
          } catch (err) {
            showNotification('❌ Copy failed. Please copy manually.');
          }
          window.getSelection().removeAllRanges(); // Deselect
        }

        function downloadFile() {
           showNotification('💾 Requesting file download...');
           google.script.run
             .withSuccessHandler(onDownloadSuccess)
             .withFailureHandler(onDownloadError)
             .downloadTexFile();
        }

        function onDownloadSuccess(result) {
            if (result && result.success) {
               showNotification('✅ File saved to Google Drive!');
               // Optionally provide link: console.log(result.url);
            } else {
               showNotification('⚠️ Download started, check Google Drive.');
            }
        }

        function onDownloadError(error) {
            showNotification('❌ Download failed: ' + error.message);
        }

        let notificationTimeout;
        function showNotification(message) {
          const notification = document.getElementById('notification');
          notification.textContent = message;
          notification.classList.add('show');

          clearTimeout(notificationTimeout); // Clear previous timeout if any
          notificationTimeout = setTimeout(() => {
            notification.classList.remove('show');
          }, 3000); // Hide after 3 seconds
        }
      </script>
    </body>
    </html>`;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(900)
    .setHeight(700);

  SlidesApp.getUi().showModalDialog(htmlOutput, '📄 Nagwa LaTeX Export');
}


function downloadTexFile() {
  const props = PropertiesService.getDocumentProperties();
  const latex = props.getProperty('latest_latex');

  if (!latex) {
    // Return an error object instead of throwing
    return { success: false, error: 'No LaTeX code found in properties. Please export first.' };
    // throw new Error('No LaTeX code found. Please export first.');
  }

  const presentation = SlidesApp.getActivePresentation();
  const presentationName = presentation.getName() || 'Untitled_Presentation'; // Handle untitled presentations
  // Sanitize filename further: replace spaces and invalid chars
  const filename = presentationName.replace(/[^\w\s-]/gi, '').replace(/\s+/g, '_') + '.tex';

  try {
      const file = DriveApp.createFile(filename, latex, MimeType.PLAIN_TEXT);

      // No UI alert here, return success and URL for the HTML dialog callback
      Logger.log('File saved: ' + file.getUrl());
      return {
        success: true,
        url: file.getUrl(),
        filename: filename
      };
   } catch (e) {
       Logger.log('Error saving file to Drive: ' + e);
       // Return error object
       return { success: false, error: 'Could not save file to Google Drive: ' + e.message };
       // throw new Error('Could not save file to Google Drive: ' + e.message);
   }
}

// ══════════════════════════════════════════════════════
// HELP
// ══════════════════════════════════════════════════════

function showHelp() {
  const ui = SlidesApp.getUi();

  // Updated help text reflecting the new Handwriting workflow
  const helpText = `
    📚 **Nagwa LaTeX Converter - Help**

    ═══════════════════════════════════

    ✍️ **HANDWRITING (Recommended Method):**
    1. Menu → Handwriting → Open Panel
    2. Enter desired Text and Font Size
    3. Click "➕ Add Handwriting"
       ↳ A yellow placeholder box appears on the slide.
    4. Move & Rotate the yellow box to the exact desired position.
    5. The script reads the position at export time.

    ═══════════════════════════════════

    🎨 **COLOR-BASED ANNOTATIONS:**
    • Select text → Change color:
      🔴 Red = Circle
      🔵 Blue = Box
      🟢 Green = Underline
      🟡 Yellow BG = Highlight (Underline)
    • Optional: Add explanations in Speaker Notes:
        Word: explanation here
    • Menu → Annotations → Auto-Detect from Colors
      ↳ Saves annotations to Speaker Notes.

    ═══════════════════════════════════

    📐 **SHAPE-BASED ANNOTATIONS:**
    • Insert → Shape → Circle/Rectangle
    • Draw around text.
    • Format: **Transparent Fill**, Visible Border (2pt+)
    • Menu → Annotations → Auto-Detect from Shapes
      ↳ Saves annotations to Speaker Notes.

    ═══════════════════════════════════

    🖊️ **MANUAL ANNOTATIONS (Notes):**
    • Open Speaker Notes for the slide.
    • Add lines like:
        [annotation:circle:01:Text To Circle]
        [annotation:box:02:Another Text]
    • IDs (01, 02) should be unique per slide for annotations.

    ═══════════════════════════════════

    🚀 **EXPORT:**
    • Menu → Export to LaTeX
    • Review generated code.
    • Copy or Download .tex file.
    • Compile with Nagwa LaTeX environment.

    ═══════════════════════════════════

    🗑️ **CLEARING DATA:**
    • Clear Auto-Detected: Removes only sections added by color/shape detection from notes.
    • Clear All Handwriting: Removes all handwriting data (markers + notes).
  `;

  // Use HTML for better formatting in the dialog
  const htmlOutput = HtmlService.createHtmlOutput('<pre style="white-space: pre-wrap; word-wrap: break-word;">' + helpText + '</pre>')
      .setWidth(500)
      .setHeight(550);
  ui.showModalDialog(htmlOutput, '📖 Help & Workflows');
}
