import win32com.client as win32
from dotenv import load_dotenv
import time
import os

# Constants
SEND_INTERVAL = 5
RED = "\033[1;31m"
GREEN = "\033[0;32m"
BLUE = "\033[1;34m"
RESET = "\033[0;0m"

# Main
if __name__ == "__main__":

    try:
        load_dotenv()
    except Exception as e:
        print(f"{RED}ERROR: Cannot read environment file{e}{RESET}")
        os._exit(1)
    
    # Initialize Outlook application
    outlook = win32.Dispatch("outlook.application")

    # Read emails
    pointer_file = open('emails.txt', 'r')
    file_text = pointer_file.read().replace('\n', '')
    pointer_file.close()
    email_list = file_text.split(';')

    # Check the curriculum file
    name_curriculum = "curriculum.pdf"
    path_curriculum = os.path.join(os.getcwd(), name_curriculum)
    if not os.path.exists(path_curriculum):
        print("Curriculum not found, please check the file name and path.")
        os._exit(1)


    # Loop for the email list 
    print(f"{BLUE}Starting email sending process...")
    for subject in email_list:
        try:
            # Create email
            email = outlook.CreateItem(0)
            
            # Init the email config and the model for subject
            email.Subject = "Candidato a Estágio em Dados ou Desenvolvimento Web"
            email.HTMLBody = f"""
            <html>
            <body>
                <p>Olá,</p>

                <p>
                Meu nome é <strong>{os.getenv("NAME")}</strong> e estou em busca de uma oportunidade como 
                <strong>estagiário</strong> ou <strong>desenvolvedor júnior</strong> nas áreas de 
                <strong>Dados</strong> ou <strong>Desenvolvimento Web</strong>.
                </p>

                <p>
                Tenho conhecimento em <strong>PHP</strong>, <strong>Python</strong>, <strong>HTML/CSS</strong> 
                e <strong>JavaScript</strong>, além de estar sempre em busca de aprendizado contínuo e de contribuir com projetos reais.
                </p>

                <p>
                Estou à disposição para entrevistas ou conversas informais, a fim de apresentar melhor minhas habilidades e objetivos profissionais.
                </p>

                <p>
                Agradeço pela atenção!
                </p>

                <p>
                Atenciosamente,<br>
                <strong>{os.getenv("NAME")}</strong><br>
                {os.getenv("LINKEDIN")}<br>
                {os.getenv("GITHUB")}<br>
                {os.getenv("PHONE")}
                </p>
            </body>
            </html>
            """

            email.To = subject
            email.Attachments.Add(path_curriculum)
            email.Send()
            print(f"{GREEN}Email sent to: {subject}")
        except Exception as e:
            print(f"{RED}ERROR sending email: {e}")
            continue

        time.sleep(SEND_INTERVAL)  

    # Reset the console color
    print(f"{BLUE}Finish email process{RESET}")

