#!/usr/bin/env python3
"""
HTML to PowerPoint Converter
Converts HTML slides from JSON to PPTX format.

Usage: python3 html_to_pptx.py input.json [output.pptx]
"""

import sys
import json
import asyncio
import tempfile
import os
from pathlib import Path
from playwright.async_api import async_playwright
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, PP_PARAGRAPH_ALIGNMENT, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
import urllib.request
import io

# Slide dimensions (1920x1080 - 16:9)
SLIDE_WIDTH = 1920
SLIDE_HEIGHT = 1080


async def extract_elements_from_html(html_content):
    """Extract text elements and background from HTML using Playwright."""
    
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page(viewport={'width': SLIDE_WIDTH, 'height': SLIDE_HEIGHT})
        
        # Create temporary HTML file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False) as f:
            f.write(html_content)
            temp_path = f.name
        
        try:
            await page.goto(f'file://{temp_path}')
            await page.wait_for_timeout(500)  # Wait for fonts and rendering
            
            # Extract elements using JavaScript
            elements = await page.evaluate("""
                () => {
                    const elements = [];
                    
                    // Helper to parse color (including gradients - extract first color)
                    const parseColor = (colorStr) => {
                        if (!colorStr || colorStr === 'transparent' || colorStr === 'rgba(0, 0, 0, 0)') return null;
                        
                        // Handle gradients by extracting the first color
                        if (colorStr.includes('gradient')) {
                            const rgbMatch = colorStr.match(/rgba?\\((\\d+),\\s*(\\d+),\\s*(\\d+)(?:,\\s*([\\d.]+))?/);
                            if (rgbMatch) {
                                const r = parseInt(rgbMatch[1]);
                                const g = parseInt(rgbMatch[2]);
                                const b = parseInt(rgbMatch[3]);
                                const alpha = rgbMatch[4] ? parseFloat(rgbMatch[4]) : 1;
                                
                                // If transparent, blend with white background to get the equivalent solid color
                                // Formula: final = alpha * color + (1 - alpha) * white
                                if (alpha < 1) {
                                    return {
                                        r: Math.round(alpha * r + (1 - alpha) * 255),
                                        g: Math.round(alpha * g + (1 - alpha) * 255),
                                        b: Math.round(alpha * b + (1 - alpha) * 255),
                                        a: 1.0
                                    };
                                }
                                
                                return {
                                    r: r,
                                    g: g,
                                    b: b,
                                    a: 1.0
                                };
                            }
                            
                            // Try hex color in gradient
                            const hexMatch = colorStr.match(/#([0-9a-fA-F]{6})/);
                            if (hexMatch) {
                                const hex = hexMatch[1];
                                return {
                                    r: parseInt(hex.substr(0, 2), 16),
                                    g: parseInt(hex.substr(2, 2), 16),
                                    b: parseInt(hex.substr(4, 2), 16),
                                    a: 1
                                };
                            }
                        }
                        
                        const rgbMatch = colorStr.match(/rgba?\\((\\d+),\\s*(\\d+),\\s*(\\d+)(?:,\\s*([\\d.]+))?/);
                        if (rgbMatch) {
                            return {
                                r: parseInt(rgbMatch[1]),
                                g: parseInt(rgbMatch[2]),
                                b: parseInt(rgbMatch[3]),
                                a: rgbMatch[4] ? parseFloat(rgbMatch[4]) : 1
                            };
                        }
                        return null;
                    };
                    
                    // Extract body background
                    const body = document.body;
                    if (body) {
                        const styles = window.getComputedStyle(body);
                        const bgColor = parseColor(styles.backgroundColor);
                        const rect = body.getBoundingClientRect();
                        
                        if (bgColor && bgColor.a > 0.1) {
                            elements.push({
                                type: 'background',
                                color: bgColor,
                                coordinates: {
                                    x: 0,
                                    y: 0,
                                    width: rect.width,
                                    height: rect.height
                                }
                            });
                        }
                    }
                    
                    // Extract shapes (backgrounds and borders from divs)
                    // Also extract text from styled elements (like pills/chips)
                    const allDivs = document.querySelectorAll('div, section, aside, header, footer, span');
                    const processedTextElements = new Set();
                    
                    allDivs.forEach(el => {
                        const styles = window.getComputedStyle(el);
                        const rect = el.getBoundingClientRect();
                        
                        // Allow small decorative elements (dots, indicators) but skip tiny ones
                        if (rect.width < 2 || rect.height < 2) return;
                        if (styles.display === 'none' || styles.visibility === 'hidden') return;
                        
                        // Check backgroundColor first, then backgroundImage (for gradients)
                        let bgColor = parseColor(styles.backgroundColor);
                        if (!bgColor || bgColor.a < 0.1) {
                            // Try background image (gradient)
                            bgColor = parseColor(styles.backgroundImage);
                        }
                        
                        const borderColor = parseColor(styles.borderTopColor);
                        const borderWidth = parseFloat(styles.borderTopWidth);
                        const borderRadius = parseFloat(styles.borderRadius);
                        
                        // Determine if this is a circle based on aspect ratio and border radius
                        const aspectRatio = rect.width / rect.height;
                        const isSquareish = aspectRatio > 0.8 && aspectRatio < 1.2;
                        const minDimension = Math.min(rect.width, rect.height);
                        const isCircle = isSquareish && borderRadius >= (minDimension / 2) * 0.9;
                        
                        // Check if this element has text and a strong background (styled chip/pill)
                        const text = (el.innerText || el.textContent).trim();
                        const hasBlockChildren = el.querySelectorAll('div, p, h1, h2, h3, h4, h5, h6, li, ul, ol').length > 0;
                        
                        // Only extract as styled_text if it's large enough (skip small badges/icons)
                        const isLargeEnough = rect.width > 60 && rect.height > 20;
                        
                        // Skip if this element is inside a semantic element (will be extracted with parent)
                        const isInsideSemanticElement = el.closest('h1, h2, h3, h4, h5, h6, p, li, button, a, label') !== null;
                        
                        if (text && !hasBlockChildren && bgColor && bgColor.a > 0.5 && borderRadius > 10 && isLargeEnough && !isInsideSemanticElement) {
                            // This is a styled chip/pill with text - create as text element with background
                            const textColor = parseColor(styles.color) || { r: 255, g: 255, b: 255, a: 1 };
                            const fontSize = parseFloat(styles.fontSize);
                            
                            let fontFamily = 'Arial';
                            if (styles.fontFamily) {
                                const fonts = styles.fontFamily.split(',').map(f => f.trim().replace(/['"]/g, ''));
                                fontFamily = fonts.find(f => !['sans-serif', 'serif', 'monospace', 'cursive', 'fantasy'].includes(f.toLowerCase())) || fonts[0];
                            }
                            
                            elements.push({
                                type: 'styled_text',
                                text: text,
                                coordinates: {
                                    x: rect.left,
                                    y: rect.top,
                                    width: rect.width,
                                    height: rect.height
                                },
                                font: {
                                    size: Math.round(fontSize * 0.75),
                                    family: fontFamily,
                                    weight: styles.fontWeight,
                                    style: styles.fontStyle
                                },
                                color: textColor,
                                alignment: styles.textAlign,
                                fill_color: bgColor,
                                border_radius: borderRadius,
                                border_color: borderColor,
                                border_width: borderWidth
                            });
                            
                            processedTextElements.add(el);
                        } else if ((bgColor && bgColor.a > 0.05) || (borderColor && borderWidth > 0)) {
                            // Skip if this element contains an image (image will be extracted separately)
                            const hasImage = el.querySelector('img') !== null;
                            if (hasImage) return;
                            
                            // Regular shape (background/border only)
                            elements.push({
                                type: 'shape',
                                coordinates: {
                                    x: rect.left,
                                    y: rect.top,
                                    width: rect.width,
                                    height: rect.height
                                },
                                fill_color: bgColor,
                                border_color: borderColor,
                                border_width: borderWidth,
                                border_radius: borderRadius,
                                is_circle: isCircle
                            });
                        }
                    });
                    
                    // Track processed elements to avoid duplication
                    const processedTableElements = new Set();
                    
                    // Extract tables with their structure
                    const tables = document.querySelectorAll('table');
                    tables.forEach(table => {
                        processedTableElements.add(table);
                        const rect = table.getBoundingClientRect();
                        if (rect.width === 0 || rect.height === 0) return;
                        
                        const styles = window.getComputedStyle(table);
                        if (styles.display === 'none' || styles.visibility === 'hidden') return;
                        
                        // Extract table data
                        const rows = [];
                        table.querySelectorAll('tr').forEach(tr => {
                            const cells = [];
                            tr.querySelectorAll('th, td').forEach(cell => {
                                const cellStyles = window.getComputedStyle(cell);
                                const cellRect = cell.getBoundingClientRect();
                                cells.push({
                                    text: cell.textContent.trim(),
                                    is_header: cell.tagName.toLowerCase() === 'th',
                                    coordinates: {
                                        x: cellRect.left,
                                        y: cellRect.top,
                                        width: cellRect.width,
                                        height: cellRect.height
                                    },
                                    alignment: cellStyles.textAlign,
                                    font_size: parseFloat(cellStyles.fontSize),
                                    font_weight: cellStyles.fontWeight,
                                    color: parseColor(cellStyles.color),
                                    bg_color: parseColor(cellStyles.backgroundColor),
                                    border_bottom_color: parseColor(cellStyles.borderBottomColor),
                                    border_bottom_width: parseFloat(cellStyles.borderBottomWidth),
                                    border_bottom_style: cellStyles.borderBottomStyle
                                });
                            });
                            if (cells.length > 0) {
                                rows.push(cells);
                            }
                        });
                        
                        if (rows.length > 0) {
                            elements.push({
                                type: 'table',
                                rows: rows,
                                coordinates: {
                                    x: rect.left,
                                    y: rect.top,
                                    width: rect.width,
                                    height: rect.height
                                }
                            });
                            
                            // Mark all table descendants as processed
                            table.querySelectorAll('*').forEach(child => processedTableElements.add(child));
                        }
                    });
                    
                    // Extract images with natural dimensions to preserve aspect ratio
                    const images = document.querySelectorAll('img');
                    images.forEach(img => {
                        const rect = img.getBoundingClientRect();
                        if (rect.width === 0 || rect.height === 0) return;
                        
                        const styles = window.getComputedStyle(img);
                        if (styles.display === 'none' || styles.visibility === 'hidden') return;
                        
                        let borderRadius = parseFloat(styles.borderRadius);
                        let isCircle = false;
                        let containerRect = rect;
                        
                        // Check if image is inside a circular container
                        const parent = img.parentElement;
                        if (parent) {
                            const parentStyles = window.getComputedStyle(parent);
                            const parentBorderRadius = parseFloat(parentStyles.borderRadius);
                            const parentRect = parent.getBoundingClientRect();
                            
                            // If parent has border-radius: 50% and is square, image should be circular
                            const parentAspectRatio = parentRect.width / parentRect.height;
                            const parentIsSquareish = parentAspectRatio > 0.8 && parentAspectRatio < 1.2;
                            const parentMinDimension = Math.min(parentRect.width, parentRect.height);
                            
                            if (parentIsSquareish && parentBorderRadius >= (parentMinDimension / 2) * 0.9) {
                                isCircle = true;
                                borderRadius = parentBorderRadius;
                                containerRect = parentRect;
                            }
                        }
                        
                        // Also check the image itself if parent wasn't circular
                        if (!isCircle) {
                            const aspectRatio = rect.width / rect.height;
                            const isSquareish = aspectRatio > 0.8 && aspectRatio < 1.2;
                            const minDimension = Math.min(rect.width, rect.height);
                            isCircle = isSquareish && borderRadius >= (minDimension / 2) * 0.9;
                        }
                        
                        elements.push({
                            type: 'image',
                            src: img.src,
                            alt: img.alt || '',
                            coordinates: {
                                x: containerRect.left,
                                y: containerRect.top,
                                width: containerRect.width,
                                height: containerRect.height
                            },
                            natural_width: img.naturalWidth,
                            natural_height: img.naturalHeight,
                            border_radius: borderRadius,
                            object_fit: styles.objectFit || 'fill',
                            is_circle: isCircle
                        });
                    });
                    
                    // Extract text elements (skip those already processed as styled_text or tables)
                    // Use semantic elements (h1-h6, p, etc) and avoid extracting from their children
                    const semanticElements = document.querySelectorAll('h1, h2, h3, h4, h5, h6, p, li, button, a, label, td, th');
                    const processedByParent = new Set();
                    
                    // First pass: Extract from semantic elements
                    semanticElements.forEach(el => {
                        // Skip if already processed
                        if (processedTextElements.has(el) || processedTableElements.has(el)) return;
                        
                        let text = (el.innerText || el.textContent).trim();
                        if (!text) return;
                        
                        // Clean up whitespace artifacts from inline badge/icon elements
                        const tagName = el.tagName.toLowerCase();
                        const hasBr = el.querySelector('br') !== null;
                        
                        if (tagName === 'p' || tagName === 'label' || tagName === 'a' || tagName === 'button') {
                            // For paragraphs and similar elements, if no explicit br tag, normalize all whitespace
                            if (!hasBr) {
                                text = text.replace(/\\s+/g, ' ');
                            } else {
                                // Has br tags, so preserve single newlines but clean up multiples
                                text = text.replace(/[ \\t]+/g, ' ').replace(/\\n\\n+/g, '\\n');
                            }
                        } else {
                            // For headings, preserve newlines from br tags
                            text = text.replace(/[ \\t]+/g, ' ').replace(/\\n\\n+/g, '\\n');
                        }
                        
                        const styles = window.getComputedStyle(el);
                        const rect = el.getBoundingClientRect();
                        
                        if (rect.width === 0 || rect.height === 0) return;
                        if (styles.display === 'none' || styles.visibility === 'hidden') return;
                        
                        const fontSize = parseFloat(styles.fontSize);
                        const textColor = parseColor(styles.color) || { r: 0, g: 0, b: 0, a: 1 };
                        
                        // Get font family
                        let fontFamily = 'Arial';
                        if (styles.fontFamily) {
                            const fonts = styles.fontFamily.split(',').map(f => f.trim().replace(/['"]/g, ''));
                            fontFamily = fonts.find(f => !['sans-serif', 'serif', 'monospace', 'cursive', 'fantasy'].includes(f.toLowerCase())) || fonts[0];
                        }
                        
                        // Check for border
                        const borderColor = parseColor(styles.borderColor || styles.borderTopColor);
                        const borderWidth = parseFloat(styles.borderWidth || styles.borderTopWidth || 0);
                        
                        elements.push({
                            type: 'text',
                            text: text,
                            coordinates: {
                                x: rect.left,
                                y: rect.top,
                                width: rect.width,
                                height: rect.height
                            },
                            font: {
                                size: Math.round(fontSize * 0.75),
                                family: fontFamily,
                                weight: styles.fontWeight,
                                style: styles.fontStyle
                            },
                            color: textColor,
                            alignment: styles.textAlign,
                            border_color: borderColor,
                            border_width: borderWidth
                        });
                        
                        // Mark this element and all descendants as processed
                        processedByParent.add(el);
                        processedTextElements.add(el);
                        el.querySelectorAll('*').forEach(child => {
                            processedByParent.add(child);
                            processedTextElements.add(child);
                        });
                    });
                    
                    // Second pass: Extract from other elements that have direct text and weren't already processed
                    const allElements = document.querySelectorAll('*');
                    
                    allElements.forEach(el => {
                        // Skip if already processed
                        if (processedTextElements.has(el) || processedTableElements.has(el) || processedByParent.has(el)) return;
                        
                        // Skip if any ancestor is a semantic element (already extracted with parent)
                        if (el.closest('h1, h2, h3, h4, h5, h6, p, li, button, a, label, td, th')) return;
                        
                        // Check if element has direct text content
                        let hasDirectText = false;
                        for (const node of el.childNodes) {
                            if (node.nodeType === Node.TEXT_NODE && node.textContent.trim().length > 0) {
                                hasDirectText = true;
                                break;
                            }
                        }
                        
                        if (!hasDirectText) return;
                        
                        const text = (el.innerText || el.textContent).trim();
                        if (!text) return;
                        
                        const styles = window.getComputedStyle(el);
                        const rect = el.getBoundingClientRect();
                        
                        if (rect.width === 0 || rect.height === 0) return;
                        if (styles.display === 'none' || styles.visibility === 'hidden') return;
                        
                        // Skip if contains block-level children
                        const blockChildren = el.querySelectorAll('div, p, h1, h2, h3, h4, h5, h6, li, ul, ol');
                        if (blockChildren.length > 0) return;
                        
                        const fontSize = parseFloat(styles.fontSize);
                        const textColor = parseColor(styles.color) || { r: 0, g: 0, b: 0, a: 1 };
                        
                        // Get font family
                        let fontFamily = 'Arial';
                        if (styles.fontFamily) {
                            const fonts = styles.fontFamily.split(',').map(f => f.trim().replace(/['"]/g, ''));
                            fontFamily = fonts.find(f => !['sans-serif', 'serif', 'monospace', 'cursive', 'fantasy'].includes(f.toLowerCase())) || fonts[0];
                        }
                        
                        // Check for border
                        const borderColor = parseColor(styles.borderColor || styles.borderTopColor);
                        const borderWidth = parseFloat(styles.borderWidth || styles.borderTopWidth || 0);
                        
                        elements.push({
                            type: 'text',
                            text: text,
                            coordinates: {
                                x: rect.left,
                                y: rect.top,
                                width: rect.width,
                                height: rect.height
                            },
                            font: {
                                size: Math.round(fontSize * 0.75),
                                family: fontFamily,
                                weight: styles.fontWeight,
                                style: styles.fontStyle
                            },
                            color: textColor,
                            alignment: styles.textAlign,
                            border_color: borderColor,
                            border_width: borderWidth
                        });
                    });
                    
                    return elements;
                }
            """)
            
        finally:
            await browser.close()
            os.unlink(temp_path)
        
        return elements


def pixels_to_inches(pixels, dpi=96):
    """Convert pixels to inches for PowerPoint."""
    return pixels / dpi


def create_slide(prs, elements):
    """Create a PowerPoint slide from extracted elements."""
    
    # Add blank slide
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    
    # Set slide background if found
    background_elem = next((e for e in elements if e['type'] == 'background'), None)
    if background_elem:
        try:
            bg_color = background_elem['color']
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(bg_color['r'], bg_color['g'], bg_color['b'])
        except:
            pass
    
    # Add shapes first (backgrounds and borders)
    for elem in elements:
        if elem['type'] != 'shape':
            continue
        
        coords = elem['coordinates']
        left = pixels_to_inches(coords['x'])
        top = pixels_to_inches(coords['y'])
        width = pixels_to_inches(coords['width'])
        height = pixels_to_inches(coords['height'])
        
        # Determine shape type
        is_circle = elem.get('is_circle', False)
        border_radius = elem.get('border_radius', 0)
        
        if is_circle:
            # Perfect circle
            shape = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                Inches(left),
                Inches(top),
                Inches(width),
                Inches(height)
            )
        elif border_radius > 20:
            # Rounded rectangle
            shape = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(left),
                Inches(top),
                Inches(width),
                Inches(height)
            )
        else:
            # Regular rectangle
            shape = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(left),
                Inches(top),
                Inches(width),
                Inches(height)
            )
        
        # Set fill color
        fill_color = elem.get('fill_color')
        if fill_color and fill_color['a'] > 0:
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(fill_color['r'], fill_color['g'], fill_color['b'])
        else:
            shape.fill.background()
        
        # Set border
        border_color = elem.get('border_color')
        border_width = elem.get('border_width', 0)
        
        if border_color and border_width > 0:
            shape.line.color.rgb = RGBColor(border_color['r'], border_color['g'], border_color['b'])
            shape.line.width = Pt(border_width)
        else:
            shape.line.fill.background()
        
        # Remove shadow
        shape.shadow.inherit = False
    
    # Add styled text elements (chips/pills with backgrounds)
    for elem in elements:
        if elem['type'] != 'styled_text':
            continue
        
        coords = elem['coordinates']
        left = pixels_to_inches(coords['x'])
        top = pixels_to_inches(coords['y'])
        width = pixels_to_inches(coords['width'])
        height = pixels_to_inches(coords['height'])
        
        # Create rounded rectangle shape
        border_radius = elem.get('border_radius', 0)
        if border_radius > 10:
            shape = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(left),
                Inches(top),
                Inches(width),
                Inches(height)
            )
            # Adjust rounding
            adjustment = min(border_radius / min(coords['width'], coords['height']), 0.5)
            try:
                shape.adjustments[0] = adjustment
            except:
                pass
        else:
            shape = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(left),
                Inches(top),
                Inches(width),
                Inches(height)
            )
        
        # Set background color
        fill_color = elem.get('fill_color')
        if fill_color and fill_color.get('a', 1) > 0:
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(fill_color['r'], fill_color['g'], fill_color['b'])
        else:
            # No fill
            shape.fill.background()
        
        # Set border if present, otherwise remove it
        border_color = elem.get('border_color')
        border_width = elem.get('border_width', 0)
        
        if border_color and border_width > 0:
            shape.line.color.rgb = RGBColor(border_color['r'], border_color['g'], border_color['b'])
            shape.line.width = Pt(border_width)
        else:
            shape.line.fill.background()
        
        # Remove shadow
        shape.shadow.inherit = False
        
        # Add text
        text_frame = shape.text_frame
        text_frame.text = elem['text']
        text_frame.word_wrap = False  # Don't wrap text in pills/chips
        text_frame.margin_left = Inches(0.1)
        text_frame.margin_right = Inches(0.1)
        text_frame.margin_top = Inches(0.05)
        text_frame.margin_bottom = Inches(0.05)
        
        # Prepare formatting for chips
        font_family = elem['font'].get('family', 'Arial')
        font_map = {
            'Proxima Nova': 'Calibri',
            'Roobert': 'Calibri',
            'system-ui': 'Calibri',
            '-apple-system': 'Calibri',
            'BlinkMacSystemFont': 'Calibri',
            'Segoe UI': 'Segoe UI',
            'Roboto': 'Calibri',
            'Arial': 'Arial'
        }
        font_name = font_map.get(font_family, 'Calibri')
        
        font_weight = str(elem['font']['weight'])
        is_bold = font_weight in ['bold', '700', '800', '900'] or (font_weight.isdigit() and int(font_weight) >= 700)
        color = elem['color']
        
        # Apply formatting to ALL paragraphs and ALL runs
        for paragraph in text_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER
            for run in paragraph.runs:
                run.font.size = Pt(elem['font']['size'])
                run.font.name = font_name
                if is_bold:
                    run.font.bold = True
                run.font.color.rgb = RGBColor(color['r'], color['g'], color['b'])
    
    # Add tables
    for elem in elements:
        if elem['type'] != 'table':
            continue
        
        rows = elem['rows']
        if not rows:
            continue
        
        # Create a text-based table representation using text boxes
        # PowerPoint tables via python-pptx are complex, so we'll use positioned text
        for row in rows:
            for cell in row:
                coords = cell['coordinates']
                left = pixels_to_inches(coords['x'])
                top = pixels_to_inches(coords['y'])
                width = pixels_to_inches(coords['width'])
                height = pixels_to_inches(coords['height'])
                
                # Add cell background if present
                bg_color = cell.get('bg_color')
                if bg_color and bg_color.get('a', 0) > 0.05:
                    bg_shape = slide.shapes.add_shape(
                        MSO_SHAPE.RECTANGLE,
                        Inches(left),
                        Inches(top),
                        Inches(width),
                        Inches(height)
                    )
                    bg_shape.fill.solid()
                    bg_shape.fill.fore_color.rgb = RGBColor(bg_color['r'], bg_color['g'], bg_color['b'])
                    bg_shape.line.fill.background()
                    # Remove shadow
                    bg_shape.shadow.inherit = False
                
                # Add bottom border if present (for table rows)
                border_bottom_color = cell.get('border_bottom_color')
                border_bottom_width = cell.get('border_bottom_width', 0)
                border_bottom_style = cell.get('border_bottom_style', 'solid')
                
                if border_bottom_color and border_bottom_width > 0:
                    # Draw a line at the bottom of the cell
                    line = slide.shapes.add_connector(
                        MSO_CONNECTOR.STRAIGHT,
                        Inches(left),
                        Inches(top + height),
                        Inches(left + width),
                        Inches(top + height)
                    )
                    line.line.color.rgb = RGBColor(border_bottom_color['r'], border_bottom_color['g'], border_bottom_color['b'])
                    line.line.width = Pt(border_bottom_width)
                    
                    # Set dash style if dotted
                    if 'dotted' in border_bottom_style or 'dashed' in border_bottom_style:
                        try:
                            line.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT if 'dotted' in border_bottom_style else MSO_LINE_DASH_STYLE.DASH
                        except:
                            pass  # If dash style not supported, use solid
                
                # Add cell text
                textbox = slide.shapes.add_textbox(
                    Inches(left),
                    Inches(top),
                    Inches(width),
                    Inches(height)
                )
                
                text_frame = textbox.text_frame
                text_frame.text = cell['text']
                text_frame.word_wrap = True
                text_frame.margin_left = Inches(0.05)
                text_frame.margin_right = Inches(0.05)
                text_frame.margin_top = Inches(0.02)
                text_frame.margin_bottom = Inches(0.02)
                
                paragraph = text_frame.paragraphs[0]
                
                # Set alignment
                alignment = cell.get('alignment', 'left')
                alignment_map = {
                    'left': PP_ALIGN.LEFT,
                    'center': PP_ALIGN.CENTER,
                    'right': PP_ALIGN.RIGHT
                }
                paragraph.alignment = alignment_map.get(alignment, PP_ALIGN.LEFT)
                
                # Format text
                if paragraph.runs:
                    run = paragraph.runs[0]
                    font_size = cell.get('font_size', 12)
                    run.font.size = Pt(int(font_size * 0.75))
                    run.font.name = 'Calibri'
                    
                    # Bold for headers
                    if cell.get('is_header') or str(cell.get('font_weight', '')).isdigit() and int(cell.get('font_weight', 400)) >= 700:
                        run.font.bold = True
                    
                    # Color
                    color = cell.get('color', {'r': 0, 'g': 0, 'b': 0})
                    if color:
                        run.font.color.rgb = RGBColor(color['r'], color['g'], color['b'])
    
    # Add images (so text renders on top)
    for elem in elements:
        if elem['type'] != 'image':
            continue
        
        coords = elem['coordinates']
        left = pixels_to_inches(coords['x'])
        top = pixels_to_inches(coords['y'])
        width = pixels_to_inches(coords['width'])
        height = pixels_to_inches(coords['height'])
        
        try:
            img_src = elem['src']
            
            # Calculate aspect ratio to avoid skewing
            natural_width = elem.get('natural_width')
            natural_height = elem.get('natural_height')
            object_fit = elem.get('object_fit', 'fill')
            
            original_width = width
            original_height = height
            original_left = left
            original_top = top
            
            if natural_width and natural_height and natural_width > 0 and natural_height > 0:
                natural_aspect = natural_width / natural_height
                
                if object_fit == 'contain':
                    # Fit image inside bounds while maintaining aspect ratio
                    display_aspect = original_width / original_height if original_height > 0 else 1
                    
                    if natural_aspect > display_aspect:
                        # Image is wider - fit to width, center vertically
                        width = original_width
                        height = width / natural_aspect
                        left = original_left
                        top = original_top + (original_height - height) / 2
                    else:
                        # Image is taller - fit to height, center horizontally
                        height = original_height
                        width = height * natural_aspect
                        left = original_left + (original_width - width) / 2
                        top = original_top
                else:
                    # Default: maintain aspect ratio by fitting to smaller dimension
                    display_aspect = original_width / original_height if original_height > 0 else 1
                    if abs(natural_aspect - display_aspect) > 0.1:
                        if natural_aspect > display_aspect:
                            height = width / natural_aspect
                        else:
                            width = height * natural_aspect
            
            is_circle = elem.get('is_circle', False)
            pic = None
            
            # Handle HTTP/HTTPS URLs
            if img_src.startswith('http'):
                with urllib.request.urlopen(img_src) as response:
                    img_data = response.read()
                    img_stream = io.BytesIO(img_data)
                    
                    pic = slide.shapes.add_picture(
                        img_stream,
                        Inches(left),
                        Inches(top),
                        width=Inches(width),
                        height=Inches(height)
                    )
            elif os.path.exists(img_src):
                # Local file
                pic = slide.shapes.add_picture(
                    img_src,
                    Inches(left),
                    Inches(top),
                    width=Inches(width),
                    height=Inches(height)
                )
            
            # Make image circular if needed by adding oval shape mask
            if pic and is_circle:
                try:
                    # Apply oval shape geometry to make it circular
                    from pptx.oxml import parse_xml
                    
                    spPr = pic._element.spPr
                    # Add preset geometry for oval/ellipse
                    prstGeom = parse_xml(
                        '<a:prstGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" prst="ellipse">'
                        '<a:avLst/>'
                        '</a:prstGeom>'
                    )
                    # Remove existing geometry if present
                    for child in list(spPr):
                        if 'Geom' in child.tag:
                            spPr.remove(child)
                    # Insert oval geometry before any other spPr children
                    spPr.insert(0, prstGeom)
                except Exception as e:
                    pass  # If circle masking fails, keep as regular rectangular image
        except Exception as e:
            print(f"  Warning: Could not add image: {e}")
    
    # Add text elements
    for elem in elements:
        if elem['type'] != 'text':
            continue
        
        coords = elem['coordinates']
        left = pixels_to_inches(coords['x'])
        top = pixels_to_inches(coords['y'])
        width = pixels_to_inches(coords['width'])
        height = pixels_to_inches(coords['height'])
        
        # Create text box
        textbox = slide.shapes.add_textbox(
            Inches(left),
            Inches(top),
            Inches(width),
            Inches(height)
        )
        
        text_frame = textbox.text_frame
        text_frame.text = elem['text']
        
        # Disable word wrap for short text (badges, names, labels)
        text_length = len(elem['text'])
        text_frame.word_wrap = text_length >= 50
        
        # Set vertical alignment - center for short text (badges), top for others
        if len(elem['text'].strip()) <= 3:
            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        else:
            text_frame.vertical_anchor = MSO_ANCHOR.TOP
        
        # Remove margins for better positioning
        text_frame.margin_left = 0
        text_frame.margin_right = 0
        text_frame.margin_top = 0
        text_frame.margin_bottom = 0
        
        # Disable auto-fit to prevent PowerPoint from resizing text
        text_frame.auto_size = MSO_AUTO_SIZE.NONE
        
        # Set alignment and format text for ALL paragraphs (important for multi-line text)
        alignment_map = {
            'left': PP_ALIGN.LEFT,
            'center': PP_ALIGN.CENTER,
            'right': PP_ALIGN.RIGHT,
            'justify': PP_ALIGN.JUSTIFY
        }
        
        # Determine alignment
        if len(elem['text'].strip()) <= 3:
            text_alignment = PP_ALIGN.CENTER
        else:
            text_alignment = alignment_map.get(elem.get('alignment', 'left'), PP_ALIGN.LEFT)
        
        # Prepare formatting attributes
        font_family = elem['font'].get('family', 'Arial')
        # Map web fonts to system fonts
        font_map = {
            'Proxima Nova': 'Calibri',
            'Roobert': 'Calibri',
            'system-ui': 'Calibri',
            '-apple-system': 'Calibri',
            'BlinkMacSystemFont': 'Calibri',
            'Segoe UI': 'Segoe UI',
            'Roboto': 'Calibri',
            'Arial': 'Arial',
            'Helvetica': 'Arial',
            'Times New Roman': 'Times New Roman',
            'Courier New': 'Courier New',
            'monospace': 'Courier New'
        }
        font_name = font_map.get(font_family, 'Calibri')
        
        font_weight = str(elem['font']['weight'])
        is_bold = font_weight in ['bold', '700', '800', '900'] or (font_weight.isdigit() and int(font_weight) >= 700)
        is_italic = elem['font'].get('style') == 'italic'
        color = elem['color']
        
        # Apply formatting to ALL paragraphs and ALL runs (critical for multi-line text!)
        for paragraph in text_frame.paragraphs:
            paragraph.alignment = text_alignment
            for run in paragraph.runs:
                run.font.size = Pt(elem['font']['size'])
                run.font.name = font_name
                if is_bold:
                    run.font.bold = True
                if is_italic:
                    run.font.italic = True
                run.font.color.rgb = RGBColor(color['r'], color['g'], color['b'])
        
        # Add border/outline if present
        border_color = elem.get('border_color')
        border_width = elem.get('border_width', 0)
        
        if border_color and border_width > 0:
            textbox.line.color.rgb = RGBColor(border_color['r'], border_color['g'], border_color['b'])
            textbox.line.width = Pt(border_width)
        else:
            # Remove border
            textbox.line.fill.background()


async def convert_json_to_pptx(json_path, output_path):
    """Convert JSON with HTML slides to PowerPoint."""
    
    # Load JSON
    with open(json_path, 'r') as f:
        slides_data = json.load(f)
    
    print(f"Processing {len(slides_data)} slides...")
    
    # Create presentation
    prs = Presentation()
    prs.slide_width = Inches(pixels_to_inches(SLIDE_WIDTH))
    prs.slide_height = Inches(pixels_to_inches(SLIDE_HEIGHT))
    
    # Process each slide
    for idx, slide_obj in enumerate(slides_data, 1):
        slide_id = slide_obj.get('id', f'slide_{idx}')
        html_content = slide_obj['html']
        
        print(f"  [{idx}/{len(slides_data)}] {slide_id}")
        
        # Extract elements
        elements = await extract_elements_from_html(html_content)
        
        # Create slide
        create_slide(prs, elements)
    
    # Save presentation
    prs.save(output_path)
    print(f"\n✓ Created: {output_path}")


async def main():
    if len(sys.argv) < 2:
        print("Usage: python3 html_to_pptx.py <json_file> [output.pptx]")
        sys.exit(1)
    
    json_file = sys.argv[1]
    
    # Default output: same name as input but with .pptx extension
    if len(sys.argv) > 2:
        output_file = sys.argv[2]
    else:
        output_file = str(Path(json_file).with_suffix('.pptx'))
    
    try:
        await convert_json_to_pptx(json_file, output_file)
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    asyncio.run(main())

