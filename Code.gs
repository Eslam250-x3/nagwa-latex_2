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
 * - Visual handwriting placement with text preview
 * - Export to Nagwa LaTeX format
 * - Download .tex file to Google Drive
 * 
 * Version: 2.0
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
const HANDWRITE_BG_COLOR = '#FFF3CD';

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
      .addItem('🎯 Open Handwriting Panel', 'showHandwritingSidebar')
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
// HANDWRITING MANAGEMENT
// ══════════════════════════════════════════════════════

function showHandwritingSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('HandwritingSidebar')
    .setTitle('✍️ Handwriting Manager')
    .setWidth(350);
  
  SlidesApp.getUi().showSidebar(html);
}

function assignHandwriting(text, size) {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  
  if (!selection) {
    throw new Error('No selection. Please select a text box first.');
  }
  
  const selectionType = selection.getSelectionType();
  
  if (selectionType !== SlidesApp.SelectionType.PAGE_ELEMENT) {
    throw new Error('Please select a text box (shape) in the slide.');
  }
  
  const elements = selection.getPageElementRange().getPageElements();
  
  if (elements.length === 0) {
    throw new Error('No element selected.');
  }
  
  if (elements.length > 1) {
    throw new Error('Please select only ONE text box at a time.');
  }
  
  const element = elements[0];
  const elementType = element.getPageElementType();
  
  if (elementType !== SlidesApp.PageElementType.SHAPE) {
    throw new Error('Selected element must be a text box (shape).');
  }
  
  const shape = element.asShape();
  
  const timestamp = new Date().getTime();
  const handwriteId = 'hw-' + timestamp;
  
  shape.getText().setText(text);
  
  const altText = HANDWRITE_MARKER + handwriteId;
  shape.setDescription(altText);
  
  try {
    const fill = shape.getFill();
    fill.setSolidFill(HANDWRITE_BG_COLOR);
  } catch (e) {
    Logger.log('Could not set background color: ' + e);
  }
  
  let currentSlide = null;
  
  const slides = presentation.getSlides();
  for (let i = 0; i < slides.length; i++) {
    const pageElements = slides[i].getPageElements();
    for (let j = 0; j < pageElements.length; j++) {
      if (pageElements[j].getObjectId() === element.getObjectId()) {
        currentSlide = slides[i];
        break;
      }
    }
    if (currentSlide) break;
  }
  
  if (!currentSlide) {
    throw new Error('Could not find slide containing the selected element.');
  }
  
  const notesShape = currentSlide.getNotesPage().getSpeakerNotesShape();
  const existingNotes = notesShape.getText().asString();
  
  const handwriteLine = '[handwrite:' + handwriteId + ':' + size + ':' + text + ']';
  
  let updatedNotes;
  if (existingNotes.trim().length === 0) {
    updatedNotes = handwriteLine;
  } else {
    updatedNotes = existingNotes + '\n' + handwriteLine;
  }
  
  notesShape.getText().setText(updatedNotes);
  
  return {
    success: true,
    id: handwriteId,
    text: text,
    size: size,
    message: 'Handwriting assigned! Text appears in box. You can now move/rotate it.'
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
  
  const element = elements[0];
  const elementType = element.getPageElementType();
  
  if (elementType !== SlidesApp.PageElementType.SHAPE) {
    return {
      hasSelection: true,
      isValid: false,
      message: 'Selected element is not a text box/shape'
    };
  }
  
  const shape = element.asShape();
  const description = shape.getDescription();
  
  let isHandwriting = false;
  let handwriteId = null;
  let currentText = shape.getText().asString();
  
  if (description && description.indexOf(HANDWRITE_MARKER) === 0) {
    isHandwriting = true;
    handwriteId = description.replace(HANDWRITE_MARKER, '');
  }
  
  return {
    hasSelection: true,
    isValid: true,
    elementId: element.getObjectId(),
    isHandwriting: isHandwriting,
    handwriteId: handwriteId,
    currentText: currentText,
    message: isHandwriting ? 
      'This box is handwriting (ID: ' + handwriteId + ')' : 
      'Ready to assign as handwriting'
  };
}

function getCurrentSlideHandwriting() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  
  let slide;
  
  if (selection && selection.getSelectionType() === SlidesApp.SelectionType.PAGE) {
    slide = selection.getCurrentPage().asSlide();
  } else {
    const slides = presentation.getSlides();
    slide = slides[0];
  }
  
  const handwritingList = [];
  const shapes = slide.getShapes();
  
  shapes.forEach(function(shape) {
    const description = shape.getDescription();
    
    if (description && description.indexOf(HANDWRITE_MARKER) === 0) {
      const handwriteId = description.replace(HANDWRITE_MARKER, '');
      
      const notes = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
      const regex = new RegExp('\\[handwrite:' + handwriteId + ':(\\d+):([^\\]]+)\\]');
      const match = notes.match(regex);
      
      if (match) {
        handwritingList.push({
          id: handwriteId,
          size: match[1],
          text: match[2],
          shapeId: shape.getObjectId()
        });
      }
    }
  });
  
  return handwritingList;
}

function deleteHandwritingById(handwriteId) {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  
  let found = false;
  
  slides.forEach(function(slide) {
    const shapes = slide.getShapes();
    
    shapes.forEach(function(shape) {
      const description = shape.getDescription();
      
      if (description && description === HANDWRITE_MARKER + handwriteId) {
        shape.setDescription('');
        shape.getText().setText('');
        
        try {
          const fill = shape.getFill();
          fill.setSolidFill('#FFFFFF');
        } catch (e) {
          Logger.log('Could not reset background: ' + e);
        }
        
        const notesShape = slide.getNotesPage().getSpeakerNotesShape();
        const notes = notesShape.getText().asString();
        
        const regex = new RegExp('\\[handwrite:' + handwriteId + ':[^\\]]+\\]\\n?', 'g');
        const updatedNotes = notes.replace(regex, '');
        
        notesShape.getText().setText(updatedNotes);
        
        found = true;
      }
    });
  });
  
  if (!found) {
    throw new Error('Handwriting with ID ' + handwriteId + ' not found.');
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
    const notes = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
    
    let slideHandwriting = [];
    
    shapes.forEach(function(shape) {
      const description = shape.getDescription();
      
      if (description && description.indexOf(HANDWRITE_MARKER) === 0) {
        const handwriteId = description.replace(HANDWRITE_MARKER, '');
        
        const regex = new RegExp('\\[handwrite:' + handwriteId + ':(\\d+):([^\\]]+)\\]');
        const match = notes.match(regex);
        
        if (match) {
          slideHandwriting.push({
            id: handwriteId,
            size: match[1],
            text: match[2]
          });
        }
      }
    });
    
    if (slideHandwriting.length > 0) {
      list += '\n📄 Slide ' + (slideIndex + 1) + ':\n';
      
      slideHandwriting.forEach(function(hw) {
        list += '  ✍️ Size:' + hw.size + 'pt - "' + hw.text + '"\n';
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
    'This will remove all handwriting markers and data.\n\n' +
    'Text boxes will remain but will be cleared.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  
  let count = 0;
  
  slides.forEach(function(slide) {
    const shapes = slide.getShapes();
    
    shapes.forEach(function(shape) {
      const description = shape.getDescription();
      
      if (description && description.indexOf(HANDWRITE_MARKER) === 0) {
        shape.setDescription('');
        shape.getText().setText('');
        
        try {
          const fill = shape.getFill();
          fill.setSolidFill('#FFFFFF');
        } catch (e) {}
        
        count++;
      }
    });
    
    const notesShape = slide.getNotesPage().getSpeakerNotesShape();
    const notes = notesShape.getText().asString();
    
    const regex = /\[handwrite:hw-\d+:[^\]]+\]\n?/g;
    const updatedNotes = notes.replace(regex, '');
    
    notesShape.getText().setText(updatedNotes);
  });
  
  ui.alert('✅ Cleared!', 'Removed ' + count + ' handwriting elements.', ui.ButtonSet.OK);
}

// ══════════════════════════════════════════════════════
// COORDINATE CONVERSION
// ══════════════════════════════════════════════════════

function emuToCm(emu, dimension) {
  let relativePosition;
  
  if (dimension === 'x') {
    relativePosition = emu / SLIDE_WIDTH_EMU;
    return relativePosition * BEAMER_WIDTH_CM;
  } else if (dimension === 'y') {
    relativePosition = emu / SLIDE_HEIGHT_EMU;
    return (1 - relativePosition) * BEAMER_HEIGHT_CM;
  }
  
  return 0;
}

function convertRotation(slidesRotation) {
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
      saveMergedAnnotations(slide, merged);
      totalAnnotations += merged.length;
      totalNotes += merged.filter(function(a) { return a.note; }).length;
    }
  });
  
  ui.alert(
    '✅ Detection Complete!',
    'Found:\n' +
    '• ' + totalAnnotations + ' colored annotations\n' +
    '• ' + totalNotes + ' explanation notes\n\n' +
    'Check Speaker Notes and Export to LaTeX',
    ui.ButtonSet.OK
  );
}

function detectColoredText(slide) {
  const colored = [];
  const shapes = slide.getShapes();
  
  shapes.forEach(function(shape) {
    const description = shape.getDescription();
    if (description && description.indexOf(HANDWRITE_MARKER) === 0) {
      return;
    }
    
    const textRange = shape.getText();
    const text = textRange.asString();
    
    if (text.trim().length === 0) return;
    
    const runs = textRange.getRuns();
    
    runs.forEach(function(run) {
      const runText = run.asString().trim();
      if (runText.length === 0) return;
      
      const style = run.getTextStyle();
      
      const foreColor = style.getForegroundColor();
      if (foreColor && foreColor.getColorType() === SlidesApp.ColorType.RGB) {
        const rgb = foreColor.asRgbColor();
        const r = Math.round(rgb.getRed() * 255);
        const g = Math.round(rgb.getGreen() * 255);
        const b = Math.round(rgb.getBlue() * 255);
        
        let type = null;
        
        if (r > 200 && g < 100 && b < 100) {
          type = 'circle';
        } else if (r < 100 && g < 100 && b > 200) {
          type = 'box';
        } else if (r < 100 && g > 200 && b < 100) {
          type = 'underline';
        }
        
        if (type) {
          colored.push({
            text: runText,
            type: type
          });
        }
      }
      
      const bgColor = style.getBackgroundColor();
      if (bgColor && bgColor.getColorType() === SlidesApp.ColorType.RGB) {
        const rgb = bgColor.asRgbColor();
        const r = Math.round(rgb.getRed() * 255);
        const g = Math.round(rgb.getGreen() * 255);
        const b = Math.round(rgb.getBlue() * 255);
        
        if (r > 200 && g > 200 && b < 100) {
          colored.push({
            text: runText,
            type: 'underline'
          });
        }
      }
    });
  });
  
  return colored;
}

function parseNotesExplanations(slide) {
  const notes = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
  const explanations = {};
  
  const lines = notes.split('\n');
  
  lines.forEach(function(line) {
    line = line.trim();
    
    if (line.indexOf('---') === 0) return;
    if (line.indexOf('[annotation:') === 0) return;
    if (line.indexOf('[handwrite:') === 0) return;
    if (line.indexOf('💡') === 0) return;
    if (line.length === 0) return;
    
    const colonIndex = line.indexOf(':');
    if (colonIndex !== -1 && colonIndex > 0) {
      const word = line.substring(0, colonIndex).trim();
      const explanation = line.substring(colonIndex + 1).trim();
      
      if (word && explanation) {
        explanations[word.toLowerCase()] = explanation;
      }
    }
  });
  
  return explanations;
}

function mergeAnnotationsWithNotes(coloredWords, explanations) {
  const merged = [];
  
  coloredWords.forEach(function(item) {
    const word = item.text;
    const wordLower = word.toLowerCase();
    
    const annotation = {
      type: item.type,
      text: word,
      note: explanations[wordLower] || null
    };
    
    merged.push(annotation);
  });
  
  return merged;
}

function saveMergedAnnotations(slide, annotations) {
  const notesShape = slide.getNotesPage().getSpeakerNotesShape();
  const existingNotes = notesShape.getText().asString();
  
  const lines = [];
  lines.push('--- AUTO-DETECTED ANNOTATIONS ---');
  
  annotations.forEach(function(ann, index) {
    const id = (index + 1).toString().padStart(2, '0');
    
    lines.push('[annotation:' + ann.type + ':' + ann.id + ':' + ann.text + ']');
    
    if (ann.note) {
      lines.push('💡 Note: ' + ann.note);
    }
  });
  
  lines.push('--- END AUTO-DETECTED ---');
  lines.push('');
  
  const newSection = lines.join('\n');
  
  const cleanedNotes = existingNotes.replace(
    /--- AUTO-DETECTED ANNOTATIONS ---[\s\S]*?--- END AUTO-DETECTED ---\n*/g,
    ''
  );
  
  const updatedNotes = newSection + '\n' + cleanedNotes.trim();
  
  notesShape.getText().setText(updatedNotes);
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
    'This will analyze all slides and detect:\n\n' +
    '⭕ Circles → Circle annotations\n' +
    '📦 Rectangles → Box annotations\n' +
    '📏 Long rectangles → Underline\n\n' +
    'Shapes must have:\n' +
    '• Transparent/light fill\n' +
    '• Visible border\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  let processed = 0;
  
  slides.forEach(function(slide, index) {
    try {
      processSlideAnnotations(slide);
      processed++;
    } catch (error) {
      Logger.log('Error processing slide ' + (index + 1) + ': ' + error);
    }
  });
  
  ui.alert(
    '✅ Detection Complete!',
    'Processed: ' + processed + '/' + slides.length + ' slides\n\n' +
    'Check Speaker Notes for auto-generated annotations.',
    ui.ButtonSet.OK
  );
}

function processSlideAnnotations(slide) {
  const shapes = slide.getShapes();
  
  const annotationShapes = [];
  const textElements = [];
  
  shapes.forEach(function(shape) {
    const description = shape.getDescription();
    if (description && description.indexOf(HANDWRITE_MARKER) === 0) {
      return;
    }
    
    const shapeType = shape.getShapeType();
    const text = shape.getText().asString().trim();
    
    if (text.length > 0) {
      textElements.push({
        shape: shape,
        text: text,
        bounds: getElementBounds(shape),
        isList: isListElement(text)
      });
    } else {
      if (isAnnotationShape(shape)) {
        annotationShapes.push({
          shape: shape,
          type: detectAnnotationType(shape),
          bounds: getElementBounds(shape)
        });
      }
    }
  });
  
  const detectedAnnotations = [];
  let annotationCounter = 1;
  
  annotationShapes.forEach(function(annShape) {
    const matchedTexts = [];
    
    textElements.forEach(function(textEl) {
      const overlap = calculateOverlap(annShape.bounds, textEl.bounds);
      
      if (overlap > 0) {
        matchedTexts.push({
          text: textEl.text,
          overlap: overlap,
          bounds: textEl.bounds
        });
      }
    });
    
    if (matchedTexts.length > 0) {
      matchedTexts.sort(function(a, b) { return b.overlap - a.overlap; });
      
      const bestMatch = matchedTexts[0];
      const annotatedText = extractAnnotatedText(
        bestMatch.text,
        annShape.bounds,
        bestMatch.bounds
      );
      
      if (annotatedText) {
        const id = annotationCounter.toString().padStart(2, '0');
        
        detectedAnnotations.push({
          type: annShape.type,
          id: id,
          text: annotatedText
        });
        
        annotationCounter++;
      }
    }
  });
  
  if (detectedAnnotations.length > 0) {
    const notesShape = slide.getNotesPage().getSpeakerNotesShape();
    const existingNotes = notesShape.getText().asString();
    
    const autoSection = buildAutoAnnotationsSection(detectedAnnotations);
    const updatedNotes = mergeNotes(existingNotes, autoSection);
    
    notesShape.getText().setText(updatedNotes);
  }
}

function isAnnotationShape(shape) {
  const shapeType = shape.getShapeType();
  
  if (shapeType !== SlidesApp.ShapeType.ELLIPSE &&
      shapeType !== SlidesApp.ShapeType.RECTANGLE &&
      shapeType !== SlidesApp.ShapeType.ROUND_RECTANGLE) {
    return false;
  }
  
  const fill = shape.getFill();
  
  if (fill.getFillType() === SlidesApp.FillType.NONE) {
    return true;
  }
  
  if (fill.getFillType() === SlidesApp.FillType.SOLID) {
    const solidFill = fill.getSolidFill();
    const alpha = solidFill.getAlpha();
    
    if (alpha < 0.3) {
      return true;
    }
  }
  
  const line = shape.getBorder();
  if (line.getLineFill().getFillType() !== SlidesApp.FillType.NONE) {
    const lineWeight = line.getWeight();
    if (lineWeight > 1) {
      return true;
    }
  }
  
  return false;
}

function detectAnnotationType(shape) {
  const shapeType = shape.getShapeType();
  
  if (shapeType === SlidesApp.ShapeType.ELLIPSE) {
    return 'circle';
  }
  
  if (shapeType === SlidesApp.ShapeType.RECTANGLE) {
    const width = shape.getWidth();
    const height = shape.getHeight();
    const ratio = width / height;
    
    if (ratio > 3 && height < 50) {
      return 'underline';
    }
    
    return 'box';
  }
  
  return 'box';
}

function getElementBounds(element) {
  const left = element.getLeft();
  const top = element.getTop();
  const width = element.getWidth();
  const height = element.getHeight();
  
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
}

function calculateOverlap(bounds1, bounds2) {
  const overlapLeft = Math.max(bounds1.left, bounds2.left);
  const overlapRight = Math.min(bounds1.right, bounds2.right);
  const overlapTop = Math.max(bounds1.top, bounds2.top);
  const overlapBottom = Math.min(bounds1.bottom, bounds2.bottom);
  
  if (overlapLeft < overlapRight && overlapTop < overlapBottom) {
    const overlapWidth = overlapRight - overlapLeft;
    const overlapHeight = overlapBottom - overlapTop;
    return overlapWidth * overlapHeight;
  }
  
  return 0;
}

function extractAnnotatedText(fullText, annBounds, textBounds) {
  if (annBounds.width < 200) {
    const words = fullText.split(/\s+/);
    
    for (let i = 0; i < words.length; i++) {
      const word = words[i].trim();
      if (word.length > 2) {
        return word;
      }
    }
  }
  
  const lines = fullText.split('\n');
  if (lines.length > 0) {
    const firstLine = lines[0].trim();
    if (firstLine.length > 0 && firstLine.length < 100) {
      return firstLine;
    }
  }
  
  const cleaned = fullText.trim().substring(0, 50);
  return cleaned;
}

function buildAutoAnnotationsSection(annotations) {
  if (annotations.length === 0) return '';
  
  const lines = [];
  lines.push('--- AUTO-DETECTED ANNOTATIONS ---');
  
  annotations.forEach(function(ann) {
    lines.push('[annotation:' + ann.type + ':' + ann.id + ':' + ann.text + ']');
  });
  
  lines.push('--- END AUTO-DETECTED ---');
  lines.push('');
  lines.push('💡 You can edit the annotations above');
  lines.push('');
  
  return lines.join('\n');
}

function mergeNotes(existingNotes, autoSection) {
  const cleanedNotes = existingNotes.replace(
    /--- AUTO-DETECTED ANNOTATIONS ---[\s\S]*?--- END AUTO-DETECTED ---\n*/g,
    ''
  );
  
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

function addManualAnnotation(text, type) {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  
  let slide;
  
  if (selection && selection.getSelectionType() === SlidesApp.SelectionType.PAGE) {
    slide = selection.getCurrentPage().asSlide();
  } else {
    const slides = presentation.getSlides();
    slide = slides[0];
  }
  
  const notesShape = slide.getNotesPage().getSpeakerNotesShape();
  const existingNotes = notesShape.getText().asString();
  
  const existingAnnotations = parseAnnotations(existingNotes);
  let nextId = existingAnnotations.length + 1;
  const idStr = nextId.toString().padStart(2, '0');
  
  const newAnnotation = '[annotation:' + type + ':' + idStr + ':' + text + ']';
  
  let updatedNotes;
  if (existingNotes.trim().length === 0) {
    updatedNotes = newAnnotation;
  } else {
    updatedNotes = existingNotes + '\n' + newAnnotation;
  }
  
  notesShape.getText().setText(updatedNotes);
  
  return {
    success: true,
    id: idStr,
    type: type,
    text: text
  };
}

function addManualHandwriting(text) {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  
  let slide;
  
  if (selection && selection.getSelectionType() === SlidesApp.SelectionType.PAGE) {
    slide = selection.getCurrentPage().asSlide();
  } else {
    const slides = presentation.getSlides();
    slide = slides[0];
  }
  
  const notesShape = slide.getNotesPage().getSpeakerNotesShape();
  const existingNotes = notesShape.getText().asString();
  
  const timestamp = new Date().getTime();
  const idStr = 'note-' + timestamp;
  
  const newNote = '[handwrite:' + idStr + ':10:' + text + ']';
  
  let updatedNotes;
  if (existingNotes.trim().length === 0) {
    updatedNotes = newNote;
  } else {
    updatedNotes = existingNotes + '\n' + newNote;
  }
  
  notesShape.getText().setText(updatedNotes);
  
  return {
    success: true,
    id: idStr,
    text: text
  };
}

function getCurrentSlideAnnotations() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  
  let slide;
  
  if (selection && selection.getSelectionType() === SlidesApp.SelectionType.PAGE) {
    slide = selection.getCurrentPage().asSlide();
  } else {
    const slides = presentation.getSlides();
    slide = slides[0];
  }
  
  const notes = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
  
  return {
    annotations: parseAnnotations(notes),
    handwriting: parseHandwritingNotes(notes)
  };
}

function deleteAnnotation(id, type, text) {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  
  let slide;
  
  if (selection && selection.getSelectionType() === SlidesApp.SelectionType.PAGE) {
    slide = selection.getCurrentPage().asSlide();
  } else {
    const slides = presentation.getSlides();
    slide = slides[0];
  }
  
  const notesShape = slide.getNotesPage().getSpeakerNotesShape();
  const existingNotes = notesShape.getText().asString();
  
  const escapedText = text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const pattern = '\\[annotation:' + type + ':' + id + ':' + escapedText + '\\]';
  const regex = new RegExp(pattern, 'g');
  const updatedNotes = existingNotes.replace(regex, '').trim();
  
  notesShape.getText().setText(updatedNotes);
  
  return { success: true };
}

function deleteHandwritingNote(id) {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  
  let slide;
  
  if (selection && selection.getSelectionType() === SlidesApp.SelectionType.PAGE) {
    slide = selection.getCurrentPage().asSlide();
  } else {
    const slides = presentation.getSlides();
    slide = slides[0];
  }
  
  const notesShape = slide.getNotesPage().getSpeakerNotesShape();
  const existingNotes = notesShape.getText().asString();
  
  const regex = new RegExp('\\[handwrite:' + id + ':[^\\]]+\\]', 'g');
  const updatedNotes = existingNotes.replace(regex, '').trim();
  
  notesShape.getText().setText(updatedNotes);
  
  return { success: true };
}

// ══════════════════════════════════════════════════════
// PARSING HELPERS
// ══════════════════════════════════════════════════════

function parseAnnotations(notes) {
  const annotations = [];
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

function parseHandwritingNotes(notes) {
  const handwriting = [];
  const regex = /\[handwrite:(note-\d+):(\d+):([^\]]+)\]/g;
  let match;
  
  while ((match = regex.exec(notes)) !== null) {
    handwriting.push({
      id: match[1],
      size: match[2],
      text: match[3].trim()
    });
  }
  
  return handwriting;
}

function isListElement(text) {
  return text.indexOf('\n•') !== -1 || 
         text.indexOf('\n-') !== -1 || 
         text.indexOf('\n*') !== -1 ||
         /^\s*\d+\./.test(text);
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
  let totalHandwriting = 0;
  
  slides.forEach(function(slide, index) {
    const notes = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
    const annotations = parseAnnotations(notes);
    const shapes = slide.getShapes();
    
    let slideHasHandwriting = false;
    shapes.forEach(function(shape) {
      const description = shape.getDescription();
      if (description && description.indexOf(HANDWRITE_MARKER) === 0) {
        slideHasHandwriting = true;
      }
    });
    
    if (annotations.length > 0 || slideHasHandwriting) {
      preview += '\n📄 Slide ' + (index + 1) + ':\n';
      
      annotations.forEach(function(ann) {
        preview += '  🔵 ' + ann.type + '[' + ann.id + ']: "' + ann.text + '"\n';
        totalAnnotations++;
      });
      
      if (slideHasHandwriting) {
        shapes.forEach(function(shape) {
          const description = shape.getDescription();
          if (description && description.indexOf(HANDWRITE_MARKER) === 0) {
            const text = shape.getText().asString();
            preview += '  ✏️ handwrite: "' + text + '"\n';
            totalHandwriting++;
          }
        });
      }
    }
  });
  
  if (preview === '') {
    preview = 'No annotations or handwriting captured yet.\n\nUse one of the Auto-Detect options or Manual Entry.';
  } else {
    preview = 'Total: ' + totalAnnotations + ' annotations, ' + totalHandwriting + ' handwriting\n' + preview;
  }
  
  ui.alert('📋 Captured Data Preview', preview, ui.ButtonSet.OK);
}

function clearCapturedData() {
  const ui = SlidesApp.getUi();
  const result = ui.alert(
    '⚠️ Clear Auto-Detected Data',
    'This will remove auto-detected annotations from Speaker Notes.\n\n' +
    'Manual annotations and explanations will be preserved.\n' +
    'Handwriting elements will NOT be affected.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  
  slides.forEach(function(slide) {
    const notesShape = slide.getNotesPage().getSpeakerNotesShape();
    const existingNotes = notesShape.getText().asString();
    
    const cleanedNotes = existingNotes.replace(
      /--- AUTO-DETECTED ANNOTATIONS ---[\s\S]*?--- END AUTO-DETECTED ---\n*/g,
      ''
    );
    
    notesShape.getText().setText(cleanedNotes.trim());
  });
  
  ui.alert('✅ Cleared!', 'Auto-detected data removed.', ui.ButtonSet.OK);
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
  
  ui.alert(
    '🚀 Exporting to LaTeX',
    'Processing ' + slides.length + ' slides...\n\n' +
    'This may take a moment.',
    ui.ButtonSet.OK
  );
  
  try {
    const config = getSessionConfig();
    const latex = generateNagwaLatex(presentation, slides, config);
    
    const props = PropertiesService.getDocumentProperties();
    props.setProperty('latest_latex', latex);
    props.setProperty('latest_latex_timestamp', new Date().toString());
    
    showLatexPreview(latex);
    
  } catch (error) {
    Logger.log('Export error: ' + error);
    ui.alert(
      '❌ Export Failed',
      'Error: ' + error.message + '\n\n' +
      'Check View → Logs for details.',
      ui.ButtonSet.OK
    );
  }
}

function generateNagwaLatex(presentation, slides, config) {
  const latex = [];
  
  latex.push(
    '\\documentclass[' + config.documentClass + ', ' +
    'nagwalang = ' + config.language + ', ' +
    'numerals = ' + config.numerals + ', ' +
    'directions = ' + config.directions + ']{nagwa}'
  );
  latex.push('');
  
  latex.push('\\usepackage{roughnotation}');
  latex.push('\\usepackage{tikz}');
  latex.push('\\usepackage{graphicx}');
  latex.push('\\usepackage{amsmath}');
  latex.push('\\usepackage{amssymb}');
  latex.push('');
  
  latex.push('\\begin{document}');
  
  latex.push('    \\metasessionID{' + config.sessionID + '}');
  latex.push('    \\sessioncountry{' + config.country + '}');
  latex.push('    \\subject{' + config.subject + '}');
  latex.push('    \\languageofinstruction{' + config.language + '}');
  latex.push('    \\grade{' + config.grade + '}');
  latex.push('    \\term{' + config.term + '}');
  latex.push('    \\sessiontitle{' + config.sessionTitle + '}');
  latex.push('');
  
  slides.forEach(function(slide, index) {
    const slideLatex = convertSlideToLatex(slide, index);
    slideLatex.forEach(function(line) {
      latex.push(line);
    });
    latex.push('');
  });
  
  latex.push('\\end{document}');
  
  return latex.join('\n');
}

function convertSlideToLatex(slide, slideIndex) {
  const latex = [];
  
  const title = extractSlideTitle(slide);
  const shapes = slide.getShapes();
  
  const notes = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
  const annotations = parseAnnotations(notes);
  
  const handwritingElements = [];
  const normalTextElements = [];
  
  shapes.forEach(function(shape) {
    const description = shape.getDescription();
    
    if (description && description.indexOf(HANDWRITE_MARKER) === 0) {
      const handwriteId = description.replace(HANDWRITE_MARKER, '');
      const hwData = extractHandwritingData(shape, handwriteId, notes);
      
      if (hwData) {
        handwritingElements.push(hwData);
      }
    } else {
      const text = shape.getText().asString().trim();
      if (text.length > 0) {
        normalTextElements.push({
          text: text,
          isList: isListElement(text)
        });
      }
    }
  });
  
  latex.push('    % Slide ' + (slideIndex + 1) + ': ' + escapeLatex(title));
  
  latex.push('    \\slidestandalone{' + escapeLatex(title) + '}{%');
  
  if (normalTextElements.length > 0) {
    const contentLatex = generateSlideContent(normalTextElements, annotations);
    contentLatex.forEach(function(line) {
      latex.push('        ' + line);
    });
  }
  
  if (handwritingElements.length > 0) {
    latex.push('');
    latex.push('        % Handwriting elements');
    
    handwritingElements.forEach(function(hw, index) {
      const hwId = (index + 1).toString().padStart(2, '0');
      
      const hwCommand = '        \\handwrite[' + hwId + ',' + 
                       hw.x + ',' + hw.y + ',' + 
                       hw.rotation + ',' + hw.size + ']{' + 
                       escapeLatex(hw.text) + '}';
      
      latex.push(hwCommand);
    });
  }
  
  latex.push('    }');
  
  return latex;
}

function extractHandwritingData(shape, handwriteId, notes) {
  const transform = shape.getTransform();
  
  const xEmu = transform.getTranslateX();
  const yEmu = transform.getTranslateY();
  
  const xCm = emuToCm(xEmu, 'x');
  const yCm = emuToCm(yEmu, 'y');
  
  const rotationDegrees = transform.getRotation() || 0;
  const latexRotation = convertRotation(rotationDegrees);
  
  const regex = new RegExp('\\[handwrite:' + handwriteId + ':(\\d+):([^\\]]+)\\]');
  const match = notes.match(regex);
  
  if (!match) {
    Logger.log('Warning: Handwriting data not found for ID: ' + handwriteId);
    return null;
  }
  
  const size = match[1];
  const text = match[2];
  
  return {
    id: handwriteId,
    x: xCm.toFixed(2),
    y: yCm.toFixed(2),
    rotation: latexRotation.toFixed(1),
    size: size,
    text: text
  };
}

function generateSlideContent(textElements, annotations) {
  const content = [];
  
  const hasLists = textElements.some(function(el) { return el.isList; });
  
  if (hasLists) {
    content.push('\\begin{itemize}');
    
    textElements.forEach(function(element) {
      if (element.isList) {
        const items = parseListItems(element.text);
        items.forEach(function(item) {
          const annotatedItem = applyAnnotationsToText(item, annotations);
          content.push('    \\item  ' + annotatedItem);
        });
      } else {
        if (element.text.trim().length > 0) {
          const annotatedText = applyAnnotationsToText(element.text, annotations);
          content.push('    \\item  ' + annotatedText);
        }
      }
    });
    
    content.push('\\end{itemize}');
    
  } else if (textElements.length > 0) {
    textElements.forEach(function(element) {
      if (element.text.trim().length > 0) {
        const annotatedText = applyAnnotationsToText(element.text, annotations);
        content.push(escapeLatex(annotatedText));
        content.push('');
      }
    });
  }
  
  return content;
}

function parseListItems(text) {
  const lines = text.split('\n');
  const items = [];
  
  lines.forEach(function(line) {
    line = line.trim();
    if (line.length === 0) return;
    
    line = line.replace(/^[•\-\*]\s*/, '');
    line = line.replace(/^\d+\.\s*/, '');
    
    if (line.length > 0) {
      items.push(line);
    }
  });
  
  return items;
}

function applyAnnotationsToText(text, annotations) {
  let result = text;
  
  annotations.forEach(function(ann) {
    if (result.indexOf(ann.text) !== -1) {
      let annotationCommand;
      
      if (ann.type === 'circle') {
        annotationCommand = '\\annotateCircle[' + ann.id + ']{' + ann.text + '}';
      } else if (ann.type === 'cross') {
        annotationCommand = '\\annotateCross[' + ann.id + ']{' + ann.text + '}';
      } else if (ann.type === 'underline') {
        annotationCommand = '\\annotateUnderline[' + ann.id + ']{' + ann.text + '}';
      } else if (ann.type === 'box') {
        annotationCommand = '\\annotateBox[' + ann.id + ']{' + ann.text + '}';
      } else {
        annotationCommand = ann.text;
      }
      
      result = result.replace(ann.text, annotationCommand);
    }
  });
  
  return result;
}

function extractSlideTitle(slide) {
  const shapes = slide.getShapes();
  
  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i];
    const placeholderType = shape.getPlaceholderType();
    
    if (placeholderType === SlidesApp.PlaceholderType.TITLE ||
        placeholderType === SlidesApp.PlaceholderType.CENTERED_TITLE) {
      return shape.getText().asString().trim();
    }
  }
  
  for (let i = 0; i < shapes.length; i++) {
    const description = shapes[i].getDescription();
    if (description && description.indexOf(HANDWRITE_MARKER) === 0) {
      continue;
    }
    
    const text = shapes[i].getText().asString().trim();
    if (text.length > 0 && text.length < 100) {
      return text;
    }
  }
  
  return 'Untitled Slide';
}

function escapeLatex(text) {
  let result = text;
  
  result = result.replace(/\\/g, '\\textbackslash{}');
  result = result.replace(/&/g, '\\&');
  result = result.replace(/%/g, '\\%');
  result = result.replace(/\$/g, '\\$');
  result = result.replace(/#/g, '\\#');
  result = result.replace(/_/g, '\\_');
  result = result.replace(/\{/g, '\\{');
  result = result.replace(/\}/g, '\\}');
  result = result.replace(/~/g, '\\textasciitilde{}');
  result = result.replace(/\^/g, '\\textasciicircum{}');
  
  return result;
}

function showLatexPreview(latex) {
  const lines = latex.split('\n').length;
  const chars = latex.length;
  const slideMatches = latex.match(/\\slidestandalone/g);
  const slideCount = slideMatches ? slideMatches.length : 0;
  const hwMatches = latex.match(/\\handwrite\[/g);
  const hwCount = hwMatches ? hwMatches.length : 0;
  
  const htmlContent = 
    '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Arial;padding:20px;background:#f8f9fa;margin:0}' +
    '.header{background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:white;padding:20px;border-radius:8px 8px 0 0;margin:-20px -20px 20px -20px}' +
    'h2{margin:0;font-size:24px}' +
    '.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:0 0 20px 0;background:white;padding:16px;border-radius:8px}' +
    '.stat{text-align:center}' +
    '.stat-value{font-size:24px;font-weight:bold;color:#667eea}' +
    '.stat-label{font-size:11px;color:#5f6368;margin-top:4px}' +
    'textarea{width:100%;height:400px;font-family:monospace;font-size:12px;padding:12px;border:1px solid #ddd;border-radius:8px;box-sizing:border-box}' +
    'button{padding:12px 24px;margin:8px 4px 0 0;border:none;border-radius:6px;cursor:pointer;font-size:14px;font-weight:500;transition:all 0.2s}' +
    '.btn-primary{background:#667eea;color:white}' +
    '.btn-primary:hover{background:#5568d3;transform:translateY(-1px)}' +
    '.btn-secondary{background:#34a853;color:white}' +
    '.btn-secondary:hover{background:#2d9248}' +
    '</style></head><body>' +
    '<div class="header"><h2>✅ LaTeX Export Complete!</h2><div style="opacity:0.9;margin-top:8px">Ready to compile with Nagwa LaTeX</div></div>' +
    '<div class="stats">' +
    '<div class="stat"><div class="stat-value">' + slideCount + '</div><div class="stat-label">Slides</div></div>' +
    '<div class="stat"><div class="stat-value">' + hwCount + '</div><div class="stat-label">Handwriting</div></div>' +
    '<div class="stat"><div class="stat-value">' + lines.toLocaleString() + '</div><div class="stat-label">Lines</div></div>' +
    '<div class="stat"><div class="stat-value">' + chars.toLocaleString() + '</div><div class="stat-label">Characters</div></div>' +
    '</div>' +
    '<textarea id="code" readonly>' + latex.replace(/</g, '&lt;').replace(/>/g, '&gt;') + '</textarea>' +
    '<div>' +
    '<button class="btn-primary" onclick="document.getElementById(\'code\').select();document.execCommand(\'copy\');alert(\'✅ Copied to clipboard!\')">📋 Copy to Clipboard</button>' +
    '<button class="btn-secondary" onclick="google.script.run.downloadTexFile()">💾 Download .tex File</button>' +
    '<button onclick="google.script.host.close()" style="background:#5f6368;color:white">Close</button>' +
    '</div>' +
    '</body></html>';
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(900)
    .setHeight(700);
  
  SlidesApp.getUi().showModalDialog(htmlOutput, '📄 Nagwa LaTeX Export');
}

function downloadTexFile() {
  const props = PropertiesService.getDocumentProperties();
  const latex = props.getProperty('latest_latex');
  
  if (!latex) {
    throw new Error('No LaTeX code found. Please export first.');
  }
  
  const presentation = SlidesApp.getActivePresentation();
  const presentationName = presentation.getName();
  const filename = presentationName.replace(/[^a-z0-9]/gi, '_') + '.tex';
  
  const file = DriveApp.createFile(filename, latex, MimeType.PLAIN_TEXT);
  
  SlidesApp.getUi().alert(
    '✅ File Saved!',
    'File: ' + filename + '\n\n' +
    'Location: Google Drive (My Drive)\n\n' +
    'URL: ' + file.getUrl(),
    SlidesApp.getUi().ButtonSet.OK
  );
  
  return {
    success: true,
    url: file.getUrl(),
    filename: filename
  };
}

// ══════════════════════════════════════════════════════
// HELP
// ══════════════════════════════════════════════════════

function showHelp() {
  const ui = SlidesApp.getUi();
  
  const helpText = 
    '📚 Nagwa LaTeX Converter - Help\n\n' +
    '═══════════════════════════════════\n\n' +
    '✍️ HANDWRITING:\n' +
    '1. Insert → Text box\n' +
    '2. Move & rotate to desired position\n' +
    '3. Menu → Handwriting → Open Panel\n' +
    '4. Enter text and size\n' +
    '5. Click "Assign as Handwriting"\n' +
    '6. Text appears in box (yellow background)\n' +
    '7. You can still move/rotate it!\n\n' +
    '═══════════════════════════════════\n\n' +
    '🎨 COLOR-BASED ANNOTATIONS:\n' +
    '• Select text and change color:\n' +
    '  🔴 Red = Circle annotation\n' +
    '  🔵 Blue = Box annotation\n' +
    '  🟢 Green = Underline annotation\n' +
    '  🟡 Yellow background = Highlight\n\n' +
    '• In Speaker Notes, add explanations:\n' +
    '  Word: explanation here\n\n' +
    '• Menu → Auto-Detect from Colors\n\n' +
    '═══════════════════════════════════\n\n' +
    '📐 SHAPE-BASED ANNOTATIONS:\n' +
    '• Insert → Shape → Circle/Rectangle\n' +
    '• Draw around text you want to annotate\n' +
    '• Format: No fill, visible border (2pt+)\n' +
    '• Menu → Auto-Detect from Shapes\n\n' +
    '═══════════════════════════════════\n\n' +
    '🖊️ MANUAL ENTRY:\n' +
    '• Menu → Manual Entry (Sidebar)\n' +
    '• Type text and select annotation type\n' +
    '• Click Add Annotation\n\n' +
    '═══════════════════════════════════\n\n' +
    '🚀 EXPORT:\n' +
    '• Menu → Export to LaTeX\n' +
    '• Copy code or download .tex file\n' +
    '• Compile with Nagwa LaTeX\n\n';
  
  ui.alert('📖 Help', helpText, ui.ButtonSet.OK);
}
