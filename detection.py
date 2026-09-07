"""
SpecCleanse Detection Module

Handles detection of removable content in DOCX specification documents.
Each detector class handles a specific content type.
"""

import re
from dataclasses import dataclass, field
from enum import Enum
from typing import Optional
from lxml import etree

from docx_xml import W, W_NS, compile_patterns, toggle_on


class ContentType(Enum):
    """Types of content that can be detected for removal."""
    SPECIFIER_NOTE = "specifier_note"
    COPYRIGHT = "copyright"
    HIDDEN_TEXT = "hidden_text"
    SPECAGENT = "specagent"
    EDITORIAL_ARTIFACT = "editorial_artifact"
    PRESERVE = "preserve"  # Content that should NOT be removed


@dataclass
class Detection:
    """Represents a detected piece of removable content."""
    content_type: ContentType
    element: etree._Element
    text: str
    confidence: float  # 0.0 to 1.0
    reason: str
    parent_paragraph: Optional[etree._Element] = None
    
    def __repr__(self):
        preview = self.text[:50] + "..." if len(self.text) > 50 else self.text
        return f"Detection({self.content_type.value}, '{preview}', conf={self.confidence:.2f})"


@dataclass
class PatternConfig:
    """Configuration for a pattern-based detector.

    Patterns arrive already compiled: every regex in ``patterns.yaml`` is
    compiled when the engine is built, so a bad pattern surfaces once, with
    its section and index, instead of once per file as "Processing error".
    """
    enabled: bool = True
    text_patterns: list[re.Pattern] = field(default_factory=list)
    low_confidence_patterns: list[re.Pattern] = field(default_factory=list)
    formatting_signals: dict = field(default_factory=dict)
    style_names: list[str] = field(default_factory=list)


class BaseDetector:
    """Base class for content detectors."""
    
    content_type: ContentType
    
    def __init__(self, config: PatternConfig):
        self.config = config
        self.compiled_patterns = config.text_patterns
        self.compiled_low_confidence_patterns = config.low_confidence_patterns

    def detect(self, element: etree._Element, text: str) -> Optional[Detection]:
        """
        Detect if element should be removed.
        Returns Detection if content should be removed, None otherwise.
        """
        raise NotImplementedError
    
    def _get_run_formatting(self, run: etree._Element) -> dict:
        """Extract formatting properties from a run."""
        formatting = {
            "italic": False,
            "bold": False,
            "color": None,
            "highlight": None,
            "hidden": False,
            "style": None,
        }
        
        rpr = run.find(f"{W}rPr")
        if rpr is None:
            return formatting

        # Toggle properties are on when present without w:val, and OFF when
        # w:val is 0/false/off.  Word writes <w:vanish w:val="0"/> to un-hide
        # a run that inherits hidden from its style, so presence is not truth.
        formatting["italic"] = toggle_on(rpr, f"{W}i")
        formatting["bold"] = toggle_on(rpr, f"{W}b")
        formatting["hidden"] = toggle_on(rpr, f"{W}vanish")

        
        # Check color
        color_elem = rpr.find(f"{W}color")
        if color_elem is not None:
            formatting["color"] = color_elem.get(f"{W}val")
            
        # Check highlight
        highlight_elem = rpr.find(f"{W}highlight")
        if highlight_elem is not None:
            formatting["highlight"] = highlight_elem.get(f"{W}val")
            
        # Check character style
        style_elem = rpr.find(f"{W}rStyle")
        if style_elem is not None:
            formatting["style"] = style_elem.get(f"{W}val")
            
        return formatting
    
    def _get_paragraph_style(self, para: etree._Element) -> Optional[str]:
        """Get paragraph style name."""
        ppr = para.find(f"{W}pPr")
        if ppr is None:
            return None
        style_elem = ppr.find(f"{W}pStyle")
        if style_elem is None:
            return None
        return style_elem.get(f"{W}val")


class SpecifierNoteDetector(BaseDetector):
    """Detects specifier notes and editorial comments."""
    
    content_type = ContentType.SPECIFIER_NOTE
    
    def detect(self, element: etree._Element, text: str) -> Optional[Detection]:
        if not self.config.enabled or not text.strip():
            return None
            
        confidence = 0.0
        reasons = []
        
        # Check text patterns
        for pattern in self.compiled_patterns:
            if pattern.search(text):
                confidence += 0.6
                reasons.append(f"Pattern match: {pattern.pattern}")
                break
        
        # Check formatting signals
        if element.tag == f"{W}r":
            formatting = self._get_run_formatting(element)
            
            # Italic text with color is a strong signal
            if formatting["italic"]:
                confidence += 0.2
                reasons.append("Italic text")
                
            # Red/blue text is a strong signal
            fmt_colors = self.config.formatting_signals.get("colors", [])
            if formatting["color"] and formatting["color"].upper() in [c.upper() for c in fmt_colors]:
                confidence += 0.3
                reasons.append(f"Color: {formatting['color']}")
                
            # Check character style
            style_names = self.config.style_names or []
            if formatting["style"] and formatting["style"] in style_names:
                confidence += 0.8
                reasons.append(f"Style: {formatting['style']}")
        
        # Check paragraph style
        elif element.tag == f"{W}p":
            para_style = self._get_paragraph_style(element)
            style_names = self.config.style_names or []
            if para_style and para_style in style_names:
                confidence += 0.8
                reasons.append(f"Paragraph style: {para_style}")
        
        if confidence >= 0.5:
            return Detection(
                content_type=self.content_type,
                element=element,
                text=text,
                confidence=min(confidence, 1.0),
                reason="; ".join(reasons)
            )
        
        return None


class CopyrightDetector(BaseDetector):
    """Detects copyright notices."""
    
    content_type = ContentType.COPYRIGHT
    
    def detect(self, element: etree._Element, text: str) -> Optional[Detection]:
        if not self.config.enabled or not text.strip():
            return None
        
        confidence = 0.0
        reasons = []
        
        # Check patterns
        for pattern in self.compiled_patterns:
            if pattern.search(text):
                confidence += 0.7
                reasons.append(f"Pattern match: {pattern.pattern}")
        
        # Multiple copyright indicators = high confidence
        matches = sum(1 for p in self.compiled_patterns if p.search(text))
        if matches >= 2:
            confidence += 0.2
            reasons.append(f"Multiple indicators: {matches}")
        
        if confidence >= 0.5:
            return Detection(
                content_type=self.content_type,
                element=element,
                text=text,
                confidence=min(confidence, 1.0),
                reason="; ".join(reasons)
            )
        
        return None


class HiddenTextDetector(BaseDetector):
    """Detects hidden text (vanish property)."""
    
    content_type = ContentType.HIDDEN_TEXT
    
    def detect(self, element: etree._Element, text: str) -> Optional[Detection]:
        if not self.config.enabled or not text.strip():
            # A run with no text carries structure, not content: a field
            # character, a footnote reference, a picture.  Removing it because
            # the run happens to be marked hidden breaks whatever it anchors.
            return None

        # Check for vanish property in run
        if element.tag == f"{W}r":
            formatting = self._get_run_formatting(element)
            if formatting["hidden"]:
                return Detection(
                    content_type=self.content_type,
                    element=element,
                    text=text,
                    confidence=1.0,
                    reason="Hidden text (vanish property)"
                )
        
        return None


class SpecAgentDetector(BaseDetector):
    """Detects SpecAgent.com references."""
    
    content_type = ContentType.SPECAGENT
    
    def detect(self, element: etree._Element, text: str) -> Optional[Detection]:
        if not self.config.enabled or not text.strip():
            return None
        
        for pattern in self.compiled_patterns:
            if pattern.search(text):
                return Detection(
                    content_type=self.content_type,
                    element=element,
                    text=text,
                    confidence=1.0,
                    reason=f"SpecAgent reference: {pattern.pattern}"
                )
        
        return None


class EditorialArtifactDetector(BaseDetector):
    """Detects editorial artifacts like placeholders and instructions."""
    
    content_type = ContentType.EDITORIAL_ARTIFACT
    
    def detect(self, element: etree._Element, text: str) -> Optional[Detection]:
        if not self.config.enabled or not text.strip():
            return None

        for pattern in self.compiled_patterns:
            if pattern.search(text):
                return Detection(
                    content_type=self.content_type,
                    element=element,
                    text=text,
                    confidence=0.8,
                    reason=f"Editorial artifact (high-confidence): {pattern.pattern}"
                )

        confidence = 0.0
        reasons = []

        matched_low_confidence_pattern = None
        for pattern in self.compiled_low_confidence_patterns:
            if pattern.search(text):
                matched_low_confidence_pattern = pattern.pattern
                confidence += 0.3
                reasons.append(f"Low-confidence pattern: {pattern.pattern}")
                break

        if matched_low_confidence_pattern is None:
            return None

        style_names = self.config.style_names or []
        fmt_colors = [
            c.upper()
            for c in self.config.formatting_signals.get("colors", [])
        ]

        if element.tag == f"{W}r":
            formatting = self._get_run_formatting(element)
            if formatting["italic"]:
                confidence += 0.2
                reasons.append("Italic text")
            if formatting["color"] and formatting["color"].upper() in fmt_colors:
                confidence += 0.3
                reasons.append(f"Color: {formatting['color']}")
            if formatting["style"] and formatting["style"] in style_names:
                confidence += 0.8
                reasons.append(f"Style: {formatting['style']}")

        elif element.tag == f"{W}p":
            para_style = self._get_paragraph_style(element)
            if para_style and para_style in style_names:
                confidence += 0.8
                reasons.append(f"Paragraph style: {para_style}")

        if confidence >= 0.5:
            return Detection(
                content_type=self.content_type,
                element=element,
                text=text,
                confidence=min(confidence, 1.0),
                reason="; ".join(reasons),
            )

        return None


class PreserveDetector(BaseDetector):
    """Detects content that should NEVER be removed (whitelist)."""
    
    content_type = ContentType.PRESERVE
    
    def detect(self, element: etree._Element, text: str) -> Optional[Detection]:
        if not self.config.enabled or not text.strip():
            return None
        
        for pattern in self.compiled_patterns:
            if pattern.search(text):
                return Detection(
                    content_type=self.content_type,
                    element=element,
                    text=text,
                    confidence=1.0,
                    reason=f"Preserved content: {pattern.pattern}"
                )
        
        return None


class DetectionEngine:
    """
    Main detection engine that coordinates all detectors.
    """
    
    def __init__(self, config: dict):
        """Initialize with configuration dictionary (from YAML).

        Every regex in the configuration is compiled here, so an invalid
        pattern raises ``ValueError`` naming its section and index once, at
        startup, rather than failing every document with "Processing error".
        """
        self.config = config
        self.detectors = self._create_detectors()
        self.preserve_detector = PreserveDetector(
            self._make_pattern_config(
                config.get("preserve_patterns", {}), "preserve_patterns"
            )
        )

    def _make_pattern_config(self, section: dict, section_name: str) -> PatternConfig:
        """Create PatternConfig from config section, compiling its patterns."""
        return PatternConfig(
            enabled=section.get("enabled", True),
            text_patterns=compile_patterns(
                section.get("text_patterns", []), f"{section_name}.text_patterns"
            ),
            low_confidence_patterns=compile_patterns(
                section.get("low_confidence_patterns", []),
                f"{section_name}.low_confidence_patterns",
            ),
            formatting_signals=section.get("formatting_signals", {}),
            style_names=(
                section.get("paragraph_styles", []) + 
                section.get("character_styles", [])
            )
        )
    
    def _create_detectors(self) -> list[BaseDetector]:
        """Create detector instances from config."""
        detectors = []
        
        # Map config sections to detector classes
        detector_map = {
            "specifier_notes": SpecifierNoteDetector,
            "copyright_notices": CopyrightDetector,
            "hidden_text": HiddenTextDetector,
            "specagent_references": SpecAgentDetector,
            "editorial_artifacts": EditorialArtifactDetector,
        }
        
        # Add style-based config to specifier notes
        style_config = self.config.get("style_based_detection", {})
        
        for section_name, detector_class in detector_map.items():
            section = self.config.get(section_name, {})
            config = self._make_pattern_config(section, section_name)
            
            # Add style names for detectors that use editorial style signals.
            if section_name in {"specifier_notes", "editorial_artifacts"}:
                config.style_names = (
                    style_config.get("paragraph_styles", []) +
                    style_config.get("character_styles", [])
                )
            if section_name == "editorial_artifacts":
                specifier_fmt = self.config.get("specifier_notes", {}).get("formatting_signals", {})
                if not config.formatting_signals:
                    config.formatting_signals = specifier_fmt
            
            detectors.append(detector_class(config))
        
        return detectors
    
    def detect_in_element(self, element: etree._Element, text: str) -> list[Detection]:
        """
        Run all detectors on an element.
        Returns list of detections (may be empty).
        """
        # First check if content should be preserved
        preserve = self.preserve_detector.detect(element, text)
        if preserve:
            return [preserve]
        
        # Run all removal detectors
        detections = []
        for detector in self.detectors:
            detection = detector.detect(element, text)
            if detection:
                detections.append(detection)
        
        return detections
    
    def should_remove(self, detections: list[Detection]) -> bool:
        """
        Determine if content should be removed based on detections.
        Returns False if any PRESERVE detection exists.
        """
        for d in detections:
            if d.content_type == ContentType.PRESERVE:
                return False
        
        # Remove if any detection with confidence >= 0.5
        return any(d.confidence >= 0.5 for d in detections)
