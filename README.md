# genai_generate_cv_template
Run the project:
1. Install the packages from requirements.txt
        pip install -r requirements.txt

2. Add OPENAI_API_KEY in the .env file :
    Instructions of how to get OPENAI_API_KEY is here : 
     https://help.openai.com/en/articles/4936850-where-do-i-find-my-openai-api-key

3. Keep all the CVs need modification in the input folder. 

4. CV template is picked up from the template folder, modify if any new template is needed. This template is an input in the UI.

5. Run the application :
    python src/generateCV_UI.py

6. New CVs based on template are generated in the output folder or whatever folder user input in the UI. The folder will be created if it is not present already. 

UI of the app :

![Screenshot 2025-04-22 at 12 29 38](https://github.com/user-attachments/assets/a7186f30-3952-4dcb-8af5-ef65ec8d47f1)
