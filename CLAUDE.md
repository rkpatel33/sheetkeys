# Code style guidelines

- **Use helper functions** - Don't inline everything, use appropriate helper functions without getting too nested.
- **Use Proper Doc strings and Typing** - Document your functions, classes, and modules with clear, concise docstrings that explain their purpose, parameters, and return values.
- **Use Proper Typing** - Use Python and typescript types throughout your codebase to clarify expected input and output types. This improves code readability, helps catch bugs early, and enables better tooling support.
- **DRY (Don't Repeat Yourself)** - Write code once and reuse it rather than copying and pasting. If you find yourself writing similar code multiple times, extract it into a function, class, or module.
- **Separation of Concerns** - Each part of your code should handle one specific responsibility. Keep business logic separate from UI code, database operations separate from validation logic, etc.
- **Don't Hard Code Values** - Use constants, configuration files, or environment variables instead of embedding literal values directly in your code. This makes changes easier and reduces errors.
- **Single Responsibility Principle** - Each function, class, or module should do one thing well. If you can't easily describe what a function does in a single sentence, it's probably doing too much.
- **Keep It Simple (KISS)** - Write code that's as simple as possible while still solving the problem. Avoid unnecessary complexity, clever tricks, or over-engineering.
- **Use Meaningful Names** - Variables, functions, and classes should have descriptive names that clearly indicate their purpose. `getUserById()` is much better than `getData()`.
- **Fail Fast** - Validate inputs early and throw errors immediately when something is wrong, rather than allowing invalid data to propagate through your system.
- **Write Self-Documenting Code** - Good code should be readable without extensive comments. When comments are needed, explain *why* something is done, not *what* is being done.
- **Test Your Code** - Write tests to verify your code works correctly and to catch regressions when you make changes.
- **Use helper functions** - Don't inline everything, use appropriate helper functions without getting too nested.
- **Write Good doc strings** - Every function, class, and module should have a doc string that clearly describes its purpose, parameters, return values, and any important details or side effects. Follow standard conventions (such as Google, NumPy, or reStructuredText style) for consistency.
- **Use comments to explain code segments** - Place concise comments above complex or non-obvious code to clarify *why* a particular approach is taken, highlight edge cases, and document design decisions. Focus on intent and reasoning, not just describing the code's actions. This makes business logic self-documenting, aids future maintainability, and aligns with CLAUDE.md by ensuring comments are clear, relevant, and positioned above intricate code sections.

# Digital Design Principles

## Core Philosophy
- **Function over form** - Usability always comes first
- **Dense information architecture** - Maximize content density while maintaining readability
- **Component-based thinking** - Build reusable, consistent patterns
- **Mobile-first responsive** - Design for smallest screens, scale up
- **Performance-conscious** - Fast load times affect user experience

## Color Strategy
- **Systematic palettes** - Primary, secondary, neutral, semantic (error/success/warning)
- **Token-based** - Use color scales (50-900) for consistent variations
- **Accessibility first** - Maintain 4.5:1 contrast minimum
- **Dark mode ready** - Design with both themes in mind from start

## Typography
- **Two fonts max** - One for headings, one for body (or just one)
- **Relative units** - Use rem/em for scalability
- **Readable line lengths** - 45-75 characters per line
- **Clear hierarchy** - Size, weight, and spacing create visual order

## Layout Principles
- **Grid foundation** - Consistent columns and gutters
- **8px spacing system** - Use multiples of 8 for all spacing
- **Proximity grouping** - Related elements stay together
- **Responsive breakpoints** - Mobile (<768px), Tablet (768-1024px), Desktop (1024px+)

## Interactive Elements
- **44px minimum touch targets** - Ensures mobile usability
- **Clear state changes** - Hover, active, disabled, loading states
- **Immediate feedback** - Every interaction gets visual response
- **Keyboard navigable** - Tab order and focus indicators

## Information Density
- **Compact spacing** - Tighter padding for data-heavy interfaces
- **Table-based layouts** - For complex data presentation
- **Progressive disclosure** - Show details on demand
- **Scannable content** - Clear headers, bullet points, visual hierarchy

## Accessibility Non-negotiables
- **Semantic HTML** - Proper heading structure
- **Keyboard support** - All interactive elements reachable
- **Screen reader compatible** - ARIA labels where needed
- **Color-independent** - Don't rely solely on color for meaning

## Performance Guidelines
- **Lazy load images** - Load only what's visible
- **System fonts first** - Faster than web fonts
- **CSS over JavaScript** - For animations and interactions
- **Critical CSS inline** - Render above-fold content fast

## Design Tokens
Define once, use everywhere:
- Colors (primary, secondary, neutral scales)
- Typography (font families, sizes, line heights)
- Spacing (xs, sm, md, lg, xl)
- Borders, shadows, transitions

## Quality Checks
- Works on all major browsers
- Passes WCAG 2.1 AA
- Loads in under 3 seconds
- Functions without JavaScript
- Readable at 200% zoom