# SSLSCAN - GUI

# WINDOWS VERSION

Execute .exe file

Sslscan allow you to perform a certificates analysis on a target domain. There are 2 way to use the software :

1) sslscan (1 thread)
2) sslscan-multithreading

The second allow you to execute a scan on a large number of hosts faster than first because use multithreading. 

The software generate 2 file :

1) json
2) excel 

In the Excel file you can find all info about certificates ( like grade, expiry and weak keys ).

## Usage

sslscangui.exe

You can decide to generate excel from json or perform a scan ( 1 host or a list of host saved in txt file ). The .txt file must be in the same directory of sslscan.py.

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License
[MIT](https://github.com/jhacker91/sslscan/blob/master/License.txt)

## NB
Use only to scan your Server. Scanning other servers whitout authorization is a crime.

## Images

Schermata Home\n
<img src="ph/Schermata 2020-06-07 alle 19.41.15.png" width="550" height="380">\n
Scansione Host Singolo. Bisogna inserire il nome dell'host da scansionare\n
<img src="ph/Schermata 2020-06-07 alle 19.41.22.png" width="550" height="380">\n
Scansione lista di host. Bisogna inserire un file .txt come nel file d'esempio (1 host per riga)\n
<img src="ph/Schermata 2020-06-07 alle 19.41.30.png" width="550" height="380">\n
Analisi host. Se si ha un file json già pronto si può ripere l'analisi per generare un excel\n
<img src="ph/Schermata 2020-06-07 alle 19.41.38.png" width="550" height="380">\n
Esempio di scansione lista\n
<img src="ph/Schermata 2020-06-07 alle 19.42.46.png" width="550" height="380">\n
Possibilità di salvare i risultati (compreso gli errori generati)\n
<img src="ph/Schermata 2020-06-07 alle 19.43.08.png" width="550" height="380">\n

