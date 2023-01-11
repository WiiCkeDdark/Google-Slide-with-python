from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
# ici c'est le scope qui permet de faire savoir a Google les autorisations que tu accordes a ce programme
# sur ce lien tu as toutes les autorisations possibles : https://developers.google.com/identity/protocols/oauth2/scopes#slides
SCOPES = ['hhttps://www.googleapis.com/auth/drive']

# The ID of a sample presentation.
# ID que tu trouves dans la barre de recherche quand tu as un slide ouvert entre /d/ et /edit
# ça doit ressembler à un truc comme ça : https://docs.google.com/presentation/d/{presentationID}/edit

PRESENTATION_ID = '1P2ZKJf6YubIeItelKwfZbEL3iUS8vu98EvhNd2LFqdk'


# -> Je te conseille de toujours utiliser ce bout de code pour récupérer les identifiants, ça permet si tu as un problème de les régénérer :
# creds = None
# if os.path.exists('token.json'):
#     creds = Credentials.from_authorized_user_file('token.json', SCOPES)
# if not creds or not creds.valid:
#     if creds and creds.expired and creds.refresh_token:
#         creds.refresh(Request())
#     else:
#         flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
#         creds = flow.run_local_server(port=0)
#     with open('token.json', 'w') as token:
#         token.write(creds.to_json())

def Example(presentation_id):
    """Shows basic usage of the Slides API.
    Prints the number of slides and elements in a sample presentation.
    """
    # ici c'est le login avec google
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        # sauvegarde pour ne pas à chaque fois refaire la manip sur le navigateur
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    # ici c'est l'interaction avec les slides
    try:
        service = build('slides', 'v1', credentials=creds)

        # Call the Slides API
        presentation = service.presentations().get(
            presentationId=presentation_id).execute()
        slides = presentation.get('slides')

        print('The presentation contains {} slides:'.format(len(slides)))
        for i, slide in enumerate(slides):
            print('- Slide #{} contains {} elements.'.format(
                i + 1, len(slide.get('pageElements'))))
    except HttpError as err:
        print(err)


def create_slide(presentation_id, page_id):
    """
    Creates the Presentation the user has access to.
    """
    # ici c'est le login avec google
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        # sauvegarde pour ne pas à chaque fois refaire la manip sur le navigateur
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    # ici c'est l'interaction avec les slides
    try:
        service = build('slides', 'v1', credentials=creds)
        # Add a slide at index 1 using the predefined
        # 'TITLE_AND_TWO_COLUMNS' layout and the ID page_id.
        requests = [
            {
                'createSlide': {
                    'objectId': page_id,
                    'insertionIndex': '1',
                    'slideLayoutReference': {
                        'predefinedLayout': 'TITLE_AND_TWO_COLUMNS'
                    }
                }
            }
        ]

        # If you wish to populate the slide with elements,
        # add element create requests here, using the page_id.

        # Execute the request.
        body = {
            'requests': requests
        }
        response = service.presentations() \
            .batchUpdate(presentationId=presentation_id, body=body).execute()
        create_slide_response = response.get('replies')[0].get('createSlide')
        print(f"Created slide with ID:"
              f"{(create_slide_response.get('objectId'))}")
    except HttpError as error:
        print(f"An error occurred: {error}")
        print("Slides not created")
        return error

    return response

# Si tu veux plus d'exemple sur la manipulation de slide, je te conseille ce lien : https://developers.google.com/slides/api/guides/presentations


if __name__ == '__main__':
    create_slide(PRESENTATION_ID,
                 "New_Slide")
