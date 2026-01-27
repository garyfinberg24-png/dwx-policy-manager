# Fluent Design Mixins Guide

This guide explains how to use the shared SCSS mixins library to maintain consistent styling across the JML solution.

## Importing the Mixins

Add this import to the top of your component SCSS file:

```scss
@import '~@/styles/fluent-mixins';
```

## Available Mixins by Category

### Container Mixins

#### `@include container-padding;`
Standard 20px padding for main containers.

```scss
.myComponent {
  @include container-padding;
}
```

#### `@include loading-container;`
Centered loading spinner container with proper spacing.

```scss
.loadingContainer {
  @include loading-container;
}
```

#### `@include empty-state;`
Centered empty state with icon and text styling.

```scss
.emptyState {
  @include empty-state;
}
```

---

### Shadow Mixins (Fluent Design Elevation)

Use these instead of custom `box-shadow` values:

```scss
.card {
  @include shadow-depth4;   // Subtle elevation
  @include shadow-depth8;   // Default card elevation
  @include shadow-depth16;  // Hover states
  @include shadow-depth64;  // Modals and flyouts
}
```

---

### Border Radius

Standard border radius values:

```scss
.element {
  @include radius-small;    // 4px - inputs, small elements
  @include radius-medium;   // 6px - cards, buttons
  @include radius-large;    // 8px - large cards, panels
  @include radius-capsule;  // 16px - badges, pills
  @include radius-circle;   // 50% - avatars, icons
}
```

---

### Card Mixins

#### `@include card;`
Standard card with hover effect.

```scss
.infoCard {
  @include card;
}
```

#### `@include card-static;`
Card without hover animation.

```scss
.staticCard {
  @include card-static;
}
```

#### `@include card-compact;`
Smaller card with less padding.

```scss
.compactCard {
  @include card-compact;
}
```

---

### Grid Layout

#### Grid Gaps

```scss
.grid {
  display: grid;
  @include grid-gap-small;   // 12px
  @include grid-gap-medium;  // 16px
  @include grid-gap-large;   // 20px
}
```

#### Responsive Grid

```scss
.gridContainer {
  @include responsive-grid(250px);  // Auto-fit columns, min 250px width
}
```

---

### Typography

```scss
h1 {
  @include heading-h1;  // 24px, weight 600
}

h2 {
  @include heading-h2;  // 20px, weight 600
}

h3 {
  @include heading-h3;  // 16px, weight 600
}

p {
  @include body-text;     // 14px, weight 400
}

.caption {
  @include caption-text;  // 12px, weight 400, secondary color
}
```

---

### Header Sections

#### Gradient Header (like Help Center)

```scss
.header {
  @include header-gradient;
  // Includes .headerIcon, .title, .subtitle/.description styling
}
```

#### Simple Header

```scss
.header {
  @include header-simple;
  // Includes h1/h2 and .description styling
}
```

---

### Buttons

```scss
.primaryButton {
  @include button-primary;
}

.secondaryButton {
  @include button-secondary;
}
```

---

### Status Badges

```scss
.statusActive {
  @include status-badge(#10893e);  // Green
}

.statusPending {
  @include status-badge(#ffa500);  // Orange
}

.statusError {
  @include status-badge(#d13438);  // Red
}
```

---

### Responsive Utilities

```scss
.myComponent {
  // Mobile styles
  @include mobile {
    flex-direction: column;
  }

  // Tablet styles
  @include tablet {
    grid-template-columns: repeat(2, 1fr);
  }

  // Desktop styles
  @include desktop {
    grid-template-columns: repeat(4, 1fr);
  }
}
```

---

### Dark Theme Support

```scss
.card {
  background: white;

  @include dark-theme {
    background: rgba(255, 255, 255, 0.05);
  }
}
```

---

### Scrollbars

```scss
.scrollableContainer {
  overflow-y: auto;
  @include custom-scrollbar;
}
```

---

### Animations

```scss
.fadeInElement {
  @include fade-in;
}

.slideUpElement {
  @include slide-up;
}
```

---

### Utility Mixins

```scss
// Flexbox centering
.centered {
  @include flex-center;
}

// Flex column
.column {
  @include flex-column;
}

// Text truncation
.truncated {
  @include text-truncate;
}

// Multi-line truncation (line clamp)
.clamped {
  @include line-clamp(3);  // Show 3 lines max
}

// Visually hidden (for accessibility)
.srOnly {
  @include visually-hidden;
}
```

---

## Migration Example

### Before (Custom Styling)

```scss
.myCard {
  padding: 20px;
  background: #ffffff;
  border-radius: 8px;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
  transition: all 0.2s ease;

  &:hover {
    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.15);
    transform: translateY(-2px);
  }
}

.title {
  font-size: 24px;
  font-weight: 600;
  color: #323130;
}

.gridContainer {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(250px, 1fr));
  gap: 20px;
}
```

### After (Using Mixins)

```scss
@import '~@/styles/fluent-mixins';

.myCard {
  @include card;
}

.title {
  @include heading-h1;
}

.gridContainer {
  @include responsive-grid(250px);
}
```

---

## Benefits

1. **Consistency** - All components use the same design tokens
2. **Maintainability** - Update styles in one place
3. **Fluent Compliance** - Follows Microsoft Fluent Design guidelines
4. **Readability** - Semantic mixin names are self-documenting
5. **Efficiency** - Less code duplication

---

## Best Practices

1. Always use mixins for shadows instead of custom `box-shadow`
2. Always use mixins for border-radius instead of custom values
3. Always use mixins for grid gaps instead of custom gap values
4. Use typography mixins for consistent font sizing
5. Import mixins at the top of your SCSS file

---

## Next Steps for Migration

Current priority files to migrate (from analysis):
- 23 files with custom box shadows → Use shadow mixins
- 22 files with non-standard border-radius → Use radius mixins
- 18 files with non-standard grid gaps → Use gap mixins

See the standardization tracking document for the full migration plan.
