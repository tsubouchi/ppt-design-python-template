# Python PowerPoint Presentation Generator

This project contains a Python script (`doer.py`) that automatically generates a PowerPoint presentation (`.pptx`) file based on predefined content and a specific design theme. The current version generates a sophisticated project proposal presentation with a minimalist black and white design using the 'Lato' font.

## Requirements

*   Python 3.x
*   `python-pptx` library

## Installation

1.  **Clone the repository (optional):**
    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```
2.  **Install the required library:**
    ```bash
    pip install python-pptx
    ```

## Usage

To generate the presentation, simply run the `doer` command from your terminal:

```bash
doer ppt
```

This will create a PowerPoint file named `project_proposal.pptx` in the same directory and automatically open it.

## Customization

You can customize the presentation by modifying the `doer.py` script:

*   **Content:** Edit the text content within the slide creation functions (e.g., `create_title_slide`, `create_executive_summary`, etc.).
*   **Design:** Adjust colors, fonts, and layouts defined in the `ColorPalette` class, font variables, and helper functions like `add_shape`, `add_background`, `apply_title_style`, `apply_body_style`.
*   **Filename:** Change the output filename in the `create_presentation` function's `prs.save()` line.

## System-wide Command Installation

To use the "doer" script as a system-wide command, copy it to a directory in your PATH (for example, /usr/local/bin) using the following commands:

```bash
sudo cp doer /usr/local/bin/doer
sudo chmod +x /usr/local/bin/doer
```

After installation, you can generate the presentation by simply running:

```bash
doer ppt
```

This command will create the presentation file "project_proposal.pptx" and automatically open it. 