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

from docx_xml import (
    W,
    StyleIndex,
    compile_patterns,
    fold_style_names,
    is_on,
    merge_spans,
    toggle_on,
)


class ContentType(Enum):
    """Types of content that can be detected for removal."""
    SPECIFIER_NOTE = "specifier_note"
    COPYRIGHT = "copyright"
    HIDDEN_TEXT = "hidden_text"
    SPECAGENT = "specagent"
    EDITORIAL_ARTIFACT = "editorial_artifact"
    INLINE_PLACEHOLDER = "inline_placeholder"  # Cut from the paragraph, not with it
    PRESERVE = "preserve"  # Content that should NOT be removed


@dataclass
class Detection:
    """Represents a detected piece of removable content."""
    content_type: ContentType
    element: etree._Element
    text: str
    confidence: float  # 0.0 to 1.0
    reason: str
    spans: list[tuple[int, int]] = field(default_factory=list)
    formatting_only: bool = False  # crossed the threshold on formatting alone
    
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
    inline_patterns: list[re.Pattern] = field(default_factory=list)
    formatting_signals: dict = field(default_factory=dict)
    style_names: list[str] = field(default_factory=list)
    formatting_only_removal: bool = False


class BaseDetector:
    """Base class for content detectors."""
    
    content_type: ContentType
    
    def __init__(self, config: PatternConfig):
        self.config = config
        self.compiled_patterns = config.text_patterns
        self.compiled_low_confidence_patterns = config.low_confidence_patterns
        self.folded_style_names = fold_style_names(config.style_names)
        self.style_index = StyleIndex()

    def bind_styles(self, style_index: StyleIndex) -> None:
        """Attach the style table of the document about to be processed."""
        self.style_index = style_index

    def detect(self, element: etree._Element, text: str) -> Optional[Detection]:
        """
        Detect if element should be removed.
        Returns Detection if content should be removed, None otherwise.
        """
        raise NotImplementedError

    def _style_matches(self, style_id: Optional[str]) -> bool:
        """True if a style, or one it is based on, is configured as editorial."""
        return self.style_index.matches(style_id, self.folded_style_names)

    def _get_run_style(self, run: etree._Element) -> Optional[str]:
        """Get the character style applied to a run."""
        rpr = run.find(f"{W}rPr")
        if rpr is None:
            return None
        style_elem = rpr.find(f"{W}rStyle")
        return style_elem.get(f"{W}val") if style_elem is not None else None

    def _get_owning_paragraph(self, element: etree._Element) -> Optional[etree._Element]:
        """Walk up from a run to the paragraph that holds it."""
        node = element.getparent()
        while node is not None:
            if node.tag == f"{W}p":
                return node
            node = node.getparent()
        return None
    
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

        # Evidence is scored in two buckets.  "Content" evidence — a text
        # pattern or an editorial style — says what the text *is*.  Formatting
        # evidence — italic, editorial colour — only says how it looks, and
        # italic + colour alone lands on exactly 0.5, the removal threshold.
        # Firms whose notes are marked only that way need it; firms whose real
        # spec text is red and italic do not, hence the switch.
        content_score = 0.0
        formatting_score = 0.0
        reasons = []

        # Check text patterns
        for pattern in self.compiled_patterns:
            if pattern.search(text):
                content_score += 0.6
                reasons.append(f"Pattern match: {pattern.pattern}")
                break
        
        # Check formatting signals
        if element.tag == f"{W}r":
            formatting = self._get_run_formatting(element)
            
            # Italic text with color is a strong signal
            if formatting["italic"]:
                formatting_score += 0.2
                reasons.append("Italic text")
                
            # Red/blue text is a strong signal
            fmt_colors = self.config.formatting_signals.get("colors", [])
            if formatting["color"] and formatting["color"].upper() in [c.upper() for c in fmt_colors]:
                formatting_score += 0.3
                reasons.append(f"Color: {formatting['color']}")
                
            # Check character style
            if self._style_matches(formatting["style"]):
                content_score += 0.8
                reasons.append(f"Style: {formatting['style']}")
        
        # Check paragraph style
        elif element.tag == f"{W}p":
            para_style = self._get_paragraph_style(element)
            if self._style_matches(para_style):
                content_score += 0.8
                reasons.append(f"Paragraph style: {para_style}")

        if content_score == 0.0 and not self.config.formatting_only_removal:
            return None

        confidence = content_score + formatting_score
        if confidence >= 0.5:
            formatting_only = content_score == 0.0
            if formatting_only:
                reasons.append("formatting-only")
            return Detection(
                content_type=self.content_type,
                element=element,
                text=text,
                confidence=min(confidence, 1.0),
                reason="; ".join(reasons),
                formatting_only=formatting_only,
            )
        
        return None


class CopyrightDetector(BaseDetector):
    """Detects copyright notices."""
    
    content_type = ContentType.COPYRIGHT
    
    def detect(self, element: etree._Element, text: str) -> Optional[Detection]:
        if not self.config.enabled or not text.strip():
            return None
        
        reasons = [
            f"Pattern match: {pattern.pattern}"
            for pattern in self.compiled_patterns
            if pattern.search(text)
        ]
        if not reasons:
            return None

        confidence = 0.7 * len(reasons)
        if len(reasons) >= 2:
            reasons.append(f"Multiple indicators: {len(reasons)}")

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
    """Detects hidden text, whether marked on the run or inherited from a style."""
    
    content_type = ContentType.HIDDEN_TEXT
    
    def detect(self, element: etree._Element, text: str) -> Optional[Detection]:
        if not self.config.enabled or not text.strip():
            # A run with no text carries structure, not content: a field
            # character, a footnote reference, a picture.  Removing it because
            # the run happens to be marked hidden breaks whatever it anchors.
            return None

        if element.tag != f"{W}r":
            return None

        reason = self._hidden_reason(element)
        if reason is None:
            return None

        return Detection(
            content_type=self.content_type,
            element=element,
            text=text,
            confidence=1.0,
            reason=reason,
        )

    def _hidden_reason(self, run: etree._Element) -> Optional[str]:
        """Why this run is hidden, or None if it is visible.

        Hidden is resolved the way Word resolves it: an explicit ``w:vanish``
        on the run wins outright — including ``w:val="0"``, which un-hides
        text that would otherwise inherit hidden — and only in its absence
        does the character style, and then the paragraph style, decide.
        """
        rpr = run.find(f"{W}rPr")
        vanish = rpr.find(f"{W}vanish") if rpr is not None else None
        if vanish is not None:
            return "Hidden text (vanish property)" if is_on(vanish) else None

        run_style = self._get_run_style(run)
        if self.style_index.is_hidden(run_style):
            return f"Hidden text (character style: {run_style})"

        para = self._get_owning_paragraph(run)
        para_style = self._get_paragraph_style(para) if para is not None else None
        if self.style_index.is_hidden(para_style):
            return f"Hidden text (paragraph style: {para_style})"

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
            return self._detect_inline(element, text)

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
            if self._style_matches(formatting["style"]):
                confidence += 0.8
                reasons.append(f"Style: {formatting['style']}")

        elif element.tag == f"{W}p":
            para_style = self._get_paragraph_style(element)
            if self._style_matches(para_style):
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

        return self._detect_inline(element, text)

    def _detect_inline(
        self, element: etree._Element, text: str
    ) -> Optional[Detection]:
        """Find placeholders sitting inside otherwise real requirement text.

        MasterSpec writes "Provide two [Verify quantity with Owner] spare
        filters per unit."  Removing that paragraph removes a requirement, so
        these matches are reported as character spans and cut out where they
        stand; only a paragraph left with nothing but placeholders is removed
        in full.  Spans are paragraph-relative, so run-level elements are not
        examined here.
        """
        if element.tag != f"{W}p" or not self.config.inline_patterns:
            return None

        spans: list[tuple[int, int]] = []
        matched: list[str] = []
        for pattern in self.config.inline_patterns:
            for match in pattern.finditer(text):
                if match.end() > match.start():
                    spans.append((match.start(), match.end()))
                    matched.append(pattern.pattern)

        if not spans:
            return None

        return Detection(
            content_type=ContentType.INLINE_PLACEHOLDER,
            element=element,
            text=text,
            confidence=0.8,
            reason="Inline placeholder: " + "; ".join(dict.fromkeys(matched)),
            spans=merge_spans(spans),
        )


class PreserveDetector(BaseDetector):
    """Detects content that should NEVER be removed (whitelist)."""
    
    content_type = ContentType.PRESERVE
    
    def detect(self, element: etree._Element, text: str) -> Optional[Detection]:
        if not self.config.enabled or not text.strip():
            return None

        # Structural styles protect headings whose text alone gives nothing
        # away: MasterSpec numbers its parts automatically, so the heading
        # "PART 1 - GENERAL" extracts as just "GENERAL".
        if element.tag == f"{W}p":
            para_style = self._get_paragraph_style(element)
            if self._style_matches(para_style):
                return Detection(
                    content_type=self.content_type,
                    element=element,
                    text=text,
                    confidence=1.0,
                    reason=f"Preserved style: {para_style}"
                )

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


# =============================================================================
# Configuration notices
# =============================================================================

#: Shipped patterns that were removed or narrowed because they deleted real
#: requirement text, mapped to what each one took.  A configuration still
#: carrying one is running the old, broader rule: ``apppaths`` prefers an
#: existing executable-adjacent or per-user ``patterns.yaml`` over the bundled
#: default, so a copy made before the change keeps it, and updating the
#: application does not update it.  Matched on the exact prior string, so an
#: edited rule is left alone rather than second-guessed.
SUPERSEDED_PATTERNS: dict[str, dict[str, str]] = {
    "copyright_notices": {
        r"may\s+not\s+be\s+reproduced":
            'removed "Shop Drawings ... may not be reproduced for use on other '
            'projects."',
        r"duplication.*?prohibited":
            'removed "...duplication of sprinkler coverage in adjacent zones is '
            'prohibited by the AHJ."',
        r"unauthorized.*?reproduction":
            'removed "Unauthorized personnel shall not have access to the fire '
            'pump room; reproduction of access keys is not permitted."',
    },
    "editorial_artifacts": {
        r"retain\s+or\s+delete":
            'removed "Provide two [retain or delete] spare filters per unit."',
        r"^\s*(?:select|choose)\s+one\b(?!-)":
            'removed "Select one of the listed manufacturers."',
    },
}


def config_notices(config: dict) -> list[str]:
    """What is worth saying about the configuration actually being used.

    Not validation — nothing here is an error, and none of it stops a run.
    These are the two things a user running an older ``patterns.yaml`` cannot
    otherwise tell: that a rule known to delete requirements is still active,
    and that removal on formatting alone is switched on.
    """
    notices: list[str] = []

    for section, superseded in SUPERSEDED_PATTERNS.items():
        active = config.get(section, {}).get("text_patterns", []) or []
        for pattern in active:
            if pattern in superseded:
                notices.append(
                    f"{section}: the pattern {pattern!r} is still active. It was "
                    f"narrowed because it {superseded[pattern]} Compare your "
                    f"patterns.yaml with the one shipped alongside this version "
                    f"to pick the change up; your edits are never overwritten."
                )

    if config.get("specifier_notes", {}).get("formatting_only_removal", False):
        notices.append(
            "specifier_notes.formatting_only_removal is on, so text is removed "
            "on italic-plus-editorial-colour alone, with no pattern or style "
            "behind it. That is off in current defaults. Preview labels such "
            "removals 'formatting-only'; tools/census_formatting reports what "
            "the setting is worth on your own documents."
        )

    return notices


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

        preserve_config = self._make_pattern_config(
            config.get("preserve_patterns", {}), "preserve_patterns"
        )
        style_config = config.get("style_based_detection", {})
        if style_config.get("enabled", True):
            preserve_config.style_names = style_config.get("preserve_styles", [])
        self.preserve_detector = PreserveDetector(preserve_config)

    def bind_styles(self, styles: dict) -> None:
        """Attach the style table of the document about to be processed.

        Called once per document, so style-based detection resolves display
        names and ``w:basedOn`` chains against that document's own styles.xml.
        """
        style_index = StyleIndex(styles)
        for detector in self.detectors:
            detector.bind_styles(style_index)
        self.preserve_detector.bind_styles(style_index)

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
            inline_patterns=compile_patterns(
                section.get("inline_patterns", []), f"{section_name}.inline_patterns"
            ),
            formatting_signals=section.get("formatting_signals", {}),
            style_names=(
                section.get("paragraph_styles", []) + 
                section.get("character_styles", [])
            ),
            formatting_only_removal=section.get("formatting_only_removal", False),
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
        style_detection_enabled = style_config.get("enabled", True)
        
        for section_name, detector_class in detector_map.items():
            section = self.config.get(section_name, {})
            config = self._make_pattern_config(section, section_name)
            
            # Add style names for detectors that use editorial style signals.
            if section_name in {"specifier_notes", "editorial_artifacts"}:
                config.style_names = (
                    (style_config.get("paragraph_styles", []) +
                     style_config.get("character_styles", []))
                    if style_detection_enabled else []
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

        Inline placeholders never remove their element: they are cut out of
        the text where they stand, so the sentence around them survives.
        """
        for d in detections:
            if d.content_type == ContentType.PRESERVE:
                return False
        
        # Remove if any detection with confidence >= 0.5
        return any(
            d.confidence >= 0.5 and d.content_type != ContentType.INLINE_PLACEHOLDER
            for d in detections
        )

    def removal_patterns(self, include_inline: bool = True) -> list[tuple[str, re.Pattern]]:
        """Every removal pattern as ``(category, compiled)``, in detector order.

        Verification classifies removals against exactly the patterns that
        caused them, so a correct low-confidence or inline removal is never
        reported as unexpected just because verification compiled a different
        list from the same file.

        ``include_inline=False`` leaves out the inline tier.  Those patterns
        never justify losing a whole paragraph — matching one only means the
        paragraph *contained* a placeholder — so whoever judges a
        whole-paragraph removal has to ask a different question of them.
        """
        patterns: list[tuple[str, re.Pattern]] = []
        for detector in self.detectors:
            if not detector.config.enabled:
                continue
            category = detector.content_type.value
            for pattern in detector.compiled_patterns:
                patterns.append((category, pattern))
            for pattern in detector.compiled_low_confidence_patterns:
                patterns.append((category, pattern))
            if include_inline:
                for pattern in detector.config.inline_patterns:
                    patterns.append((ContentType.INLINE_PLACEHOLDER.value, pattern))
        return patterns

    def inline_patterns(self) -> list[re.Pattern]:
        """Compiled inline placeholder patterns, across every enabled detector."""
        return [
            pattern
            for detector in self.detectors
            if detector.config.enabled
            for pattern in detector.config.inline_patterns
        ]

    def preserve_patterns(self) -> list[re.Pattern]:
        """Compiled preserve patterns — content that must never be removed."""
        if not self.preserve_detector.config.enabled:
            return []
        return self.preserve_detector.compiled_patterns
