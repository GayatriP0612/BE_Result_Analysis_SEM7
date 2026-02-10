import google.generativeai as genai

genai.configure(api_key="AIzaSyAGmFr7qzkyEFbSHWaeoXuIpi0imIoyUyA")

for model in genai.list_models():
    print(model.name)
