import spacy
from spacy.training.example import Example
import random

# --- 1. THE AUTO-ANNOTATOR HELPER ---
def create_training_data(text, target_entity):
    """Automatically finds the character positions of your target entity."""
    start = text.find(target_entity)
    if start == -1:
        print(f"Error: Could not find '{target_entity}' in the text.")
        return None
    end = start + len(target_entity)
    return (text, {"entities": [(start, end, "PROJECT_NAME")]})

# --- 2. YOUR TRAINING DATASET ---
# Format: create_training_data("Full sentence from PDF", "The Exact Project Name")
print("Building dataset...")
raw_data = [
    create_training_data("The primary objective of the Alpha Upgrade Initiative is to...", "Alpha Upgrade Initiative"),
    create_training_data("Project Name: NextGen Aviation System", "NextGen Aviation System"),
    create_training_data("Title: Facility Modernization 2026", "Facility Modernization 2026"),
    create_training_data("This report outlines the scope for the Automated Billing Protocol.", "Automated Billing Protocol"),
    create_training_data("Subject: Final clearance for Operation Red Sky deployment.", "Operation Red Sky"),
    create_training_data("PROJECT TITLE: EV INTEROPERABILITY STUDY", "EV INTEROPERABILITY STUDY"),
    
    # ⚠️ ADD YOUR OWN EXAMPLES HERE!
    # Copy 40 to 50 sentences from your actual DRDO PDFs and extract the names.
    # The more diverse they are (all caps, title case, different verbs), the smarter it gets!
]

# Remove any empty entries in case of typos
TRAIN_DATA = [data for data in raw_data if data is not None]

# --- 3. THE TRAINING LOOP ---
def train_spacy():
    print(f"Training on {len(TRAIN_DATA)} examples...")
    
    # Create a blank English AI model
    nlp = spacy.blank("en")
    
    # Create the Named Entity Recognition (NER) pipeline
    ner = nlp.add_pipe("ner")
    
    # Add our custom "PROJECT_NAME" label to the AI's brain
    ner.add_label("PROJECT_NAME")
    
    # Prepare the AI for training
    optimizer = nlp.begin_training()
    
    print("Beginning training loops (Epochs)...")
    # Train the model 30 times over the data
    for itn in range(30):
        random.shuffle(TRAIN_DATA)
        losses = {}
        
        for text, annotations in TRAIN_DATA:
            doc = nlp.make_doc(text)
            example = Example.from_dict(doc, annotations)
            # Update the model's weights based on your examples
            nlp.update([example], sgd=optimizer, losses=losses)
            
        print(f"Epoch {itn + 1} - Loss: {losses['ner']:.2f}")

    # Save the trained brain to a folder
    output_dir = "./custom_project_ner"
    nlp.to_disk(output_dir)
    print(f"\n✅ Model successfully trained and saved to the folder: {output_dir}")

if __name__ == "__main__":
    train_spacy()
