This is a script used  when given multiple addresses to compile a plan of action on the order of addresses to tackle based on the shortest distance from a starting point. The following points will be achieved by selecting the point in the cluster that has the shortest distance to the current one.

Using python 2.7

Git Clone Repo

Create a config.py

    import googlemaps
    gmaps = googlemaps.Client(key='Your Key')


Set Up Local Environment
Run In Terminal or git (might need to get pip, and download virtualenv):
    virtualenv venv
    source venv/bin/activate
    pip install -r requirements.txt

Add the excel file to your folder

Run In Terminal:
    python run.py


