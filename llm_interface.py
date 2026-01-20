import requests
import os
import json
from dotenv import load_dotenv

load_dotenv()  # Loads from .env into os.environ

class LLMInterface:
  def __init__(self, model="openai/gpt-4o-mini", api_key=None):
        self.url = "https://openrouter.ai/api/v1/chat/completions"
        self.model = model
        self.api_key = api_key or os.getenv("OPENAI_API_KEY")
        self.headers = {
        "Authorization": f"Bearer {self.api_key}",
        # "HTTP-Referer": "<YOUR_SITE_URL>", # Optional. Site URL for rankings on openrouter.ai.
        # "X-Title": "<YOUR_SITE_NAME>", # Optional. Site title for rankings on openrouter.ai.
      }

  def ask(self, prompt, system_prompt="You are a helpful assistant."):
    response = requests.post(
      url="https://openrouter.ai/api/v1/chat/completions",
      headers=self.headers,
      data=json.dumps({
        # "model": "qwen/qwen3-235b-a22b:free", # Optional
        "model": self.model,
        "messages": [
          {
              "role": "system",
              "content": system_prompt
          },
          {
              "role": "user",
              "content": prompt
          }
        ]
      })
    )
    if response.status_code == 200:
        response = json.loads(response.text)
        response = response["choices"][0]["message"]
        return response
    else:
        error_msg = f"❌ Error contacting the LLM API. Status: {response.status_code}"
        try:
            error_detail = response.json()
            error_msg += f", Detail: {error_detail}"
        except:
            error_msg += f", Text: {response.text[:200]}"
        print(f"\n    API Error: {error_msg}")
        return error_msg

  def check_czech_grammar(self, text):
    """
    Checks Czech grammar in the provided text.
    Returns a dict with 'has_errors', 'corrections', and 'confidence'.
    """
    system_prompt = """Jsi expert na českou gramatiku a pravopis. 
Tvým úkolem je zkontrolovat pravopis a gramatiku v českém textu.
Odpověz POUZE ve formátu JSON s těmito klíči:
- has_errors: true/false (zda text obsahuje chyby)
- corrections: seznam oprav (pokud jsou nějaké chyby), každá oprava obsahuje "original" a "corrected"
- confidence: číslo 0-100 (jak si jsi jistý/á svou kontrolou)

Příklad odpovědi:
{
  "has_errors": true,
  "corrections": [
    {"original": "Jaký je nejvetší město", "corrected": "Jaké je největší město"}
  ],
  "confidence": 95
}

Pokud není chyba, vrať:
{
  "has_errors": false,
  "corrections": [],
  "confidence": 90
}"""
    
    prompt = f"Zkontroluj prosím tento text:\n\n{text}"
    
    response = self.ask(prompt, system_prompt)
    
    # Try to parse JSON response
    try:
        if isinstance(response, dict) and 'content' in response:
            content = response['content']
            # Extract JSON from markdown code blocks if present
            if '```json' in content:
                content = content.split('```json')[1].split('```')[0].strip()
            elif '```' in content:
                content = content.split('```')[1].split('```')[0].strip()
            
            result = json.loads(content)
            return result
        else:
            return {"has_errors": False, "corrections": [], "confidence": 0, "error": "Invalid response format"}
    except json.JSONDecodeError:
        return {"has_errors": False, "corrections": [], "confidence": 0, "error": "Could not parse LLM response"}
  
  def check_answer_correctness(self, question, answer, provided_answer):
    """
    Checks if a provided answer is correct for a given question.
    Only returns positive result if LLM is confident (>= 80%).
    
    Returns a dict with 'is_correct', 'confidence', and 'explanation'.
    """
    system_prompt = """Jsi expert na kvízové otázky a fakta. 
Tvým úkolem je ověřit, zda odpověď na otázku je správná.

DŮLEŽITÉ:
- Porovnej zadanou odpověď s očekávanou odpovědí
- Buď přísný ale rozumný - menší gramatické chyby nebo překlepy nevadí
- Pokud si nejsi jistý více než 80%, vrať is_correct jako false
- Vrať POUZE JSON formát

Formát odpovědi:
{
  "is_correct": true/false,
  "confidence": číslo 0-100,
  "explanation": "krátké vysvětlení"
}"""
    
    prompt = f"""Otázka: {question}

Očekávaná odpověď: {answer}

Zadaná odpověď: {provided_answer}

Je zadaná odpověď správná?"""
    
    response = self.ask(prompt, system_prompt)
    
    # Try to parse JSON response
    try:
        if isinstance(response, str):
            # Response is error message string
            return {"is_correct": False, "confidence": 0, "explanation": response}
        
        if not isinstance(response, dict):
            return {"is_correct": False, "confidence": 0, "explanation": f"Unexpected response type: {type(response)}"}
        
        # Get content from response
        content = response.get('content', '')
        
        if not content:
            return {"is_correct": False, "confidence": 0, "explanation": "Empty response content"}
        
        # Extract JSON from markdown code blocks if present
        if '```json' in content:
            content = content.split('```json')[1].split('```')[0].strip()
        elif '```' in content:
            content = content.split('```')[1].split('```')[0].strip()
        
        # Try to parse JSON
        result = json.loads(content)
        
        # Only return positive if confidence >= 80%
        if result.get('confidence', 0) < 80:
            result['is_correct'] = False
            if 'explanation' in result:
                result['explanation'] += " (Nízká jistota LLM)"
        
        return result
    except json.JSONDecodeError as e:
        return {"is_correct": False, "confidence": 0, "explanation": f"JSON parse error: {str(e)}. Content: {content[:100]}"}
    except Exception as e:
        return {"is_correct": False, "confidence": 0, "explanation": f"Unexpected error: {str(e)}"}

