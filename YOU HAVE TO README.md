Please reach out to Dr. Pablo to request the credentials for the GitHub and Render.com accounts.

First, create a new folder on your computer and name it something memorable.

Right-click on the folder and select "Open with VS Code" to open it in Visual Studio Code.

Once VS Code is open, launch the terminal by the Ctrl + Shift + ` and run the following command to clone the project repository:

    git clone <repository-url>

Navigate to the project folder by running:

    cd advice

Install the project dependencies by running:

    pip install -r requirements.txt

To run the application, use the following command:

    python app.py

After running the app, press Ctrl + click on the localhost link to view the website in your browser.

Ensure you're signed in to VS Code with the same GitHub account.

After making changes, click the Source Control button on the left side of VS Code. Write a commit message for your changes and commit them.

Sync the changes by clicking Sync Changes.

Finally, go to Render.com and deploy your latest changes by selecting the project and choosing the Deploy your Latest commit option from the Manual Deploy dropdown and then click the
logs option on the left sidebar to see the progress and errors.
