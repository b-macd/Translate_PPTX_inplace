# Translate Powerpoints in-place

This app lets you load an Arabic language Powerpoint file and then runs the slides through a text extractor followed by an NLP model that conducts machine translation to english.

Upon completion you will be given the option to download the newly translated file.

## Running the app locally

* clone repository to your local machine
* open a terminal and navigate to the Translate_PPTX_inplace directory
* run the following commands:
  * pip install -r requirements.txt
  * streamlit run app/translator_app_external_model.py

## Running with the model on your local drive

If you want to run the same or different models on the local drive with this app, you will have to download the model and all of its config files to the same directory as this project and rename the variable `model_name` in the `translator_app_local_model_v1.py` file to match the relative filepath to the model folder.

The files for the model that I am using as the translator in this app can be located at this link: https://huggingface.co/Helsinki-NLP/opus-mt-ar-en/tree/main

