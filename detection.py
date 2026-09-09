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
    iter_own_runs,
    layout_break_offsets,
    merge_spans,
    paragraph_text,
    run_text,
    spans_cover,
    tidy_spans,
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


# =============================================================================
# Source evidence
# =============================================================================

#: Category reported for an interval that is nothing but whitespace.  Callers
#: normally filter these out before asking; the constant exists so the shape of
#: the answer never depends on which characters happened to be in the span.
WHITESPACE_ONLY = "whitespace"


@dataclass(frozen=True)
class RunEvidence:
    """One run's text, where it sits in the paragraph, and its own authority.

    ``start`` and ``end`` are offsets into the paragraph's **raw** text — the
    string :func:`docx_xml.paragraph_text` returns, before any stripping.  An
    offset computed against a trimmed or normalized string is not an offset
    into the document.
    """
    start: int
    end: int
    text: str
    category: str | None = None
    reason: str | None = None
    #: True when the detection crossed the threshold on formatting alone.
    formatting_only: bool = False

    @property
    def authorized(self) -> bool:
        """True if policy permits losing *this run's own text* — and no more."""
        return self.category is not None


@dataclass(frozen=True)
class ParagraphEvidence:
    """What the configured policy permits losing from one source paragraph.

    Evaluated against the source document alone.  Nothing the processor reports
    is consulted: what the cleaner says it did is not evidence about what the
    output contains.

    The distinction the flattened ``(category, regex)`` list could not carry is
    the one that matters here — *scope*.  ``whole_category`` says a rule
    qualified against the whole paragraph; ``runs`` say which individual runs
    carry their own authority; ``inline_spans`` say which intervals a
    placeholder rule authorizes.  A hidden run in a mixed paragraph produces one
    authorized run interval, never permission to lose the requirement beside it.
    """
    raw_text: str
    runs: tuple[RunEvidence, ...] = ()
    preserve_reason: str | None = None
    whole_category: str | None = None
    whole_reason: str | None = None
    #: True when the whole-paragraph rule crossed the threshold on formatting alone.
    whole_formatting_only: bool = False
    inline_spans: tuple[tuple[int, int], ...] = ()
    inline_reason: str | None = None
    #: False when run offsets could not be reconciled with the paragraph text,
    #: in which case interval reasoning is abandoned rather than guessed at.
    offsets_reliable: bool = True

    @property
    def preserved(self) -> bool:
        return self.preserve_reason is not None

    def authorized_spans(self) -> list[tuple[int, int]]:
        """Every interval policy permits losing, in raw-text coordinates.

        Authorized *runs* and authorized *placeholder intervals* together.  A
        whole-paragraph rule is deliberately absent: it is not an interval
        claim, and answering "may this fragment go?" with "the paragraph could
        have gone" is how a paragraph-scope match came to excuse an arbitrary
        edit inside it.
        """
        if not self.offsets_reliable:
            return []
        spans = [(r.start, r.end) for r in self.runs if r.authorized]
        spans.extend(self.inline_spans)
        return merge_spans(spans)

    def covers_all_text(self) -> bool:
        """True if the authorized intervals account for every substantive character.

        This is what licenses losing the whole paragraph in the absence of a
        whole-paragraph rule — and it is a real test rather than the presence of
        one editorial signal somewhere in the paragraph.
        """
        if not self.offsets_reliable:
            return False
        covered = bytearray(len(self.raw_text))
        for start, end in self.authorized_spans():
            for idx in range(max(0, start), min(len(covered), end)):
                covered[idx] = 1
        return all(
            covered[idx] or not char.strip()
            for idx, char in enumerate(self.raw_text)
        )

    def authorities(self) -> list[tuple[str, str, bool]]:
        """Distinct ``(category, reason, formatting_only)`` behind the intervals.

        In evidence order, so the first is the one reported when a caller needs
        a single category for a loss several rules jointly account for.
        """
        found: list[tuple[str, str, bool]] = []
        for run in self.runs:
            if run.authorized:
                found.append((run.category, run.reason or run.category, run.formatting_only))
        if self.inline_spans:
            found.append((
                ContentType.INLINE_PLACEHOLDER.value,
                self.inline_reason or ContentType.INLINE_PLACEHOLDER.value,
                False,
            ))
        return list(dict.fromkeys(found))

    def authorizes_whole_paragraph(self) -> bool:
        """True if losing this entire paragraph is something policy asked for."""
        if self.preserved:
            return False
        if self.whole_category is not None:
            return True
        return bool(self.raw_text.strip()) and self.covers_all_text()

    def reason_for_span(self, start: int, end: int) -> tuple[str, str, bool] | None:
        """The ``(category, reason, formatting_only)`` authorizing it, or None.

        Whitespace is not substantive: an authorized interval may take the
        space that fell beside it, which is what ``tidy_spans`` already does
        when the processor cuts a placeholder out.
        """
        if not self.offsets_reliable:
            return None
        first: tuple[str, str, bool] | None = None
        for idx in range(start, end):
            if not self.raw_text[idx].strip():
                continue
            hit = self._authority_at(idx)
            if hit is None:
                return None
            if first is None:
                first = hit
        return first or (WHITESPACE_ONLY, "adjacent whitespace", False)

    def _authority_at(self, idx: int) -> tuple[str, str, bool] | None:
        for run in self.runs:
            if run.authorized and run.start <= idx < run.end:
                return (run.category, run.reason or run.category, run.formatting_only)
        for start, end in self.inline_spans:
            if start <= idx < end:
                return (
                    ContentType.INLINE_PLACEHOLDER.value,
                    self.inline_reason or ContentType.INLINE_PLACEHOLDER.value,
                    False,
                )
        return None


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

    def paragraph_evidence(self, para: etree._Element) -> ParagraphEvidence:
        """What policy permits losing from ``para``, read from the source alone.

        This states the same decision the processor acts on, in the same order —
        preserve, then a whole-paragraph rule, then placeholder intervals, then
        individual runs — but as a pure read of the source document.  It never
        mutates anything and never asks the processor what it did, which is the
        boundary verification has to hold: the cleaner's account of its own work
        cannot be the evidence that the work was right.

        Run detection is skipped for a preserved or already-removable paragraph
        because the processor skips it too.  A preserved paragraph is left
        completely alone — not redacted, not trimmed — so any loss inside one is
        damage, whatever the lost text looks like.

        ``tests/test_evidence.py`` pins this against what the processor actually
        does, so the two cannot drift apart unnoticed.
        """
        raw_text = paragraph_text(para)

        preserve = self.preserve_detector.detect(para, raw_text)
        if preserve is not None:
            return ParagraphEvidence(raw_text=raw_text, preserve_reason=preserve.reason)

        para_detections = self.detect_in_element(para, raw_text)
        if self.should_remove(para_detections):
            chosen = max(
                (
                    d for d in para_detections
                    if d.confidence >= 0.5
                    and d.content_type != ContentType.INLINE_PLACEHOLDER
                ),
                key=lambda d: d.confidence,
            )
            return ParagraphEvidence(
                raw_text=raw_text,
                whole_category=chosen.content_type.value,
                whole_reason=chosen.reason,
                whole_formatting_only=chosen.formatting_only,
            )

        inline_spans, inline_reason = self._inline_evidence(para, raw_text, para_detections)

        runs: list[RunEvidence] = []
        offset = 0
        for run in iter_own_runs(para):
            text = run_text(run)
            start, offset = offset, offset + len(text)
            category = reason = None
            formatting_only = False
            if text.strip():
                run_detections = self.detect_in_element(run, text)
                if self.should_remove(run_detections):
                    best = max(
                        (
                            d for d in run_detections
                            if d.confidence >= 0.5
                            and d.content_type != ContentType.INLINE_PLACEHOLDER
                        ),
                        key=lambda d: d.confidence,
                    )
                    category, reason = best.content_type.value, best.reason
                    formatting_only = best.formatting_only
            runs.append(
                RunEvidence(start, offset, text, category, reason, formatting_only))

        return ParagraphEvidence(
            raw_text=raw_text,
            runs=tuple(runs),
            inline_spans=tuple(inline_spans),
            inline_reason=inline_reason,
            offsets_reliable=(offset == len(raw_text)),
        )

    def _inline_evidence(
        self,
        para: etree._Element,
        raw_text: str,
        detections: list[Detection],
    ) -> tuple[list[tuple[int, int]], str | None]:
        """Placeholder intervals policy authorizes cutting out of ``para``.

        A span straddling a page or column break is dropped, because the
        processor refuses to cut one: claiming authority it declines to use
        would report a paragraph it deliberately left alone as damage.
        """
        spans = [
            span
            for d in detections
            if d.content_type == ContentType.INLINE_PLACEHOLDER and d.element == para
            for span in d.spans
        ]
        if not spans:
            return [], None

        breaks = layout_break_offsets(para)
        kept = [
            span for span in tidy_spans(raw_text, merge_spans(spans))
            if not any(spans_cover([span], start, end) for start, end in breaks)
        ]
        if not kept:
            return [], None

        reasons = [
            d.reason
            for d in detections
            if d.content_type == ContentType.INLINE_PLACEHOLDER and d.element == para
        ]
        return kept, "; ".join(dict.fromkeys(reasons)) or None

    def preserve_style_names(self) -> frozenset[str]:
        """Folded names of the styles that protect a paragraph outright."""
        if not self.preserve_detector.config.enabled:
            return frozenset()
        return self.preserve_detector.folded_style_names

