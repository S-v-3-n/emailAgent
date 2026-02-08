# Outlook Email Response Agent

This agent connects to your Outlook account, fetches unread emails, deduplicates them by subject, and uses a local AI model (via Ollama) to draft replies interactively.

## Features

- **Outlook Integration**: Connects securely using OAuth2 via `msal` and `imap-tools`.
- **Local AI**: Uses `ollama` to generate privacy-focused, local AI replies.
- **Smart Deduplication**: Groups emails by subject and conversational history to avoid redundancy.
- **Context Aware**: Can load external context from a `resources/` directory to inform the AI.
- **Interactive CLI**: Reviews emails one-by-one and asks if you want to generate a draft.
- **Draft Creation**: Saves generated replies directly to your Outlook "Drafts" folder.

## Prerequisites

- **Python 3.8+**
- **Outlook Account**: An Office 365 or Outlook.com account.
- **Azure App Registration**: You need a Client ID from Azure Entra ID (formerly Azure AD) with `IMAP.AccessAsUser.All` and `SMTP.Send` permissions.
- **Ollama**: Installed and running locally (e.g., `ollama serve`).

## Installation

1. **Clone the repository** (if applicable) or navigate to the project folder.

2. **Create a virtual environment**:
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Install and Pull Ollama Model**:
   Ensure [Ollama](https://ollama.com/) is installed. Then pull the model you intend to use (default in example is `llama3.2`):
   ```bash
   ollama pull llama3.2
   ```

## Configuration

1. **Environment Variables**:
   Copy `.env.example` to `.env`:
   ```bash
   cp .env.example .env
   ```

2. **Edit `.env`**:
   Fill in your details:
   - `EMAIL_USER`: Your Outlook email address.
   - `CLIENT_ID`: Your Azure App Client ID.
   - `OLLAMA_MODEL`: The local model to use (e.g., `llama3.2`, `mistral`).
   - `TARGET_SENDER`: (Optional) Filter emails by this sender email address.

   *Note: `EMAIL_PASSWORD` is typically not used if using OAuth2, but kept for legacy support if needed.*

## Usage

Run the agent:

```bash
python main.py
```

1. The agent will authenticate (it may open a browser window for login on the first run).
2. It will fetch unread emails.
3. For each unique conversation, it will show a preview and ask: `Generate draft reply? (y/n/q)`.
4. If you select `y`, it generates a reply using the local AI and saves it to your "Drafts" folder.

## Project Structure

- `main.py`: Entry point. Handles the main loop and user interaction.
- `src/`:
  - `email_client.py`: Handles IMAP connections, fetching emails, and creating drafts.
  - `ai_engine.py`: Interfaces with Ollama to generate text.
  - `auth.py`: Handles MSAL authentication (OAuth2).
  - `config.py`: Loads environment variables.
  - `context_loader.py`: Loads text files from `resources/` for extra context.
- `resources/`: Place `.txt` or `.md` files here to give the AI more background information (e.g., project details, FAQs).

## Troubleshooting

- **Authentication Failed**: Ensure your Azure App has "Mobile and desktop applications" platform configured with `http://localhost:8080` or `http://localhost` as a Redirect URI.
- **IMAP Error**: Ensure IMAP is enabled on your Outlook account settings.
- **Ollama Connection**: Ensure `ollama` is running (`ollama list` should show your models).
