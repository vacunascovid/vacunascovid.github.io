import codecs
import datetime
import os
import requests
import locale
from pyexcel_ods import get_data
from git import Repo
from operator import itemgetter
from PIL import Image, ImageDraw, ImageFont

URL = "https://www.mscbs.gob.es/profesionales/saludPublica/ccayes/alertasActual/nCov/documentos/Informe_Comunicacion_2021"
WEB = "https://www.mscbs.gob.es/profesionales/saludPublica/ccayes/alertasActual/nCov/vacunaCovid19.htm"
locale.setlocale(locale.LC_ALL, 'esp')
POPULATION = [8464411, 1329391, 1018784, 1171543, 2175952, 582905, 2394918, 2045221, 7780479, 5057353, 1063987, 2701819,
              319914, 6779888, 1511251, 661197, 2220504, 84202, 87076, 133282, 0, 47450795]
PATH_OF_GIT_REPO = 'C:/programming/web/vacunascovid.github.io'


def update_excel():
    fecha = datetime.datetime.utcnow()

    try:
        open("{}.ods".format(fecha.strftime("%m-%d")))
        return
    except FileNotFoundError:
        while True:
            r = requests.get("{}{}.ods".format(URL, fecha.strftime("%m%d")))
            if str(r.content).find("Error 404") == -1:
                break
            fecha -= datetime.timedelta(days=1)

        with open("{}.ods".format(datetime.datetime.utcnow().strftime("%m-%d")), "wb") as file:
            file.write(r.content)


def update_html():
    date = datetime.datetime.utcnow()
    part_data = get_data("{}.ods".format(date.strftime("%m-%d")), start_row=1)
    vaccinated_list, half_vaccinated_list, percentage_list = [], [], []

    for index in range(21):
        if index == 20:
            index += 1
        pop = POPULATION[index]
        place = part_data["Comunicación"][index][0]
        doses = part_data["Comunicación"][index][5]
        used_doses = part_data["Comunicación"][index][6]
        half_vaccinated = part_data["Comunicación"][index][8]
        vaccinated_people = part_data["Comunicación"][index][9]
        dose_per = used_doses / doses * 100
        vaccinated_per = vaccinated_people / pop * 100
        half_vaccinated_per = half_vaccinated / pop * 100
        dropdown = place

        if place == "Totales":
            place = "España"
            dropdown = "Comunidades"

        if place == "C. Valenciana":
            place, dropdown = "Valencia", "Valencia"
            part_data["Comunicación"][index][0] = place

        if place in ["Asturias ", "Murcia "]:
            place = place[:-1]
            part_data["Comunicación"][index][0] = place

        html_part = """
        <div class="w3-display-container w3-center">
            <img src="img/{}.jpg" style="width:20%">
        </div>
        <div align="center" class="w3-container w3-text-black w3-margin-top">
            <h1 style="font-family: Lovelo;font-size: 50px;"> Población de {} actual: {:n} </h1>
            <br></br>
            <div class="w3-panel w3-topbar" style="width: 60%"></div>
            <h2 style="font-size: 40px;"> Personas vacunadas totalmente: {:n}</h2>
            <br></br>
            <div class="w3-light-grey w3-round w3-center" style="width:30%">
                <div class="w3-container w3-green w3-round w3-text-black w3-large" style="width:{:.2f}%;height:32px">{:.2f}%</div>
            </div>
            <br></br>
            <div class="w3-panel w3-topbar" style="width: 60%"></div>
            <h2 style="font-size: 40px;">Personas con una dosis: {:n}</h2>
            <br></br>
            <div class="w3-light-grey w3-round w3-center w3-margin-top" style="width:30%">
                <div class="w3-container w3-light-green w3-round w3-text-black w3-large" style="width:{:.2f}%;height:32px">{:.2f}%</div>
            </div>
            <br></br>
            <div class="w3-panel w3-topbar" style="width: 60%"></div>
            <h2 style="font-size: 40px;">Dosis administradas: {:n} / {:n}</h2>
            <br></br>
            <div class="w3-light-grey w3-round w3-center w3-margin-top" style="width:30%">
                <div class="w3-container w3-lime w3-round w3-text-black w3-large" style="width:{:.2f}%;height:32px">{:.2f}%</div>
            </div>
            <br></br>
            <div class="w3-panel w3-topbar" style="width: 60%"></div>
        </div>
        </div>""".format(place, place, pop, vaccinated_people, vaccinated_per, vaccinated_per, half_vaccinated,
                         half_vaccinated_per, half_vaccinated_per, used_doses, doses, dose_per, dose_per)

        html = get_html(dropdown, part_data["Comunicación"], html_part, date.strftime("%d/%m/%Y"))

        with codecs.open(PATH_OF_GIT_REPO + "/{}.html".format(place), "w", "utf-8-sig") as web:
            web.write(html)

        vaccinated_list.append((vaccinated_per, place))
        half_vaccinated_list.append((half_vaccinated_per, place))

    percentage_list = vaccinated_list.copy()
    vaccinated_list.sort(key=itemgetter(0), reverse=True)
    half_vaccinated_list.sort(key=itemgetter(0), reverse=True)

    html_part = """
        <div class="container" style="width: 100%">
            <div class="row">
                <div class="col-md-6" align="center">
                    <h1 class="w3-text-black">Vacunados totalmente</h1>
                    <div class="w3-panel w3-topbar" style="width: 80%"></div>
                    <br></br>
                    {}
                </div>
                <div class="col-md-6" align="center">
                    <h1 class="w3-text-black">Vacunados parcialmente</h1>
                    <div class="w3-panel w3-topbar" style="width: 80%"></div>
                    <br></br>
                    {}
                </div>
            </div>
        </div>""".format(get_top_list(vaccinated_list), get_top_list(half_vaccinated_list, True))

    html = get_html("Comunidades", part_data["Comunicación"], html_part, date.strftime("%d/%m/%Y"))

    with codecs.open(PATH_OF_GIT_REPO + "/index.html", "w", "utf-8-sig") as web:
        web.write(html)


def get_html(dropdown, data, html_part, date):
    html = """
    <html>
    <head>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Covid en España</title>
        <link rel="stylesheet" href="CSS/w3.css">
        <link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-alpha.6/css/bootstrap.min.css'>
        <link rel="stylesheet" href="CSS/style.css">
        <link rel="stylesheet" href="CSS/mix.css">
        <link rel="stylesheet" href="Lovelo.ttf"/>
    </head>
    <body class="crystal-clear">
        <nav class="hover-nav dropdown" style="background-color: #FFFFFF">
          <ul>
            <li><a href="España.html">España</a></li>
            <li><a href="index.html">Lista Comunidades</a></li>
            <li>
                <button class="btn btn-secondary dropdown-toggle" type="button" id="dropdownMenuButton" 
                data-toggle="dropdown" aria-haspopup="true" aria-expanded="false"> {}
                </button>
                <div class="dropdown-menu dropdown-multicol2" aria-labelledby="dropdownMenuButton">
                    {}
                </div>
            </li>
          </ul>
        </nav>
        <br></br>
        {}
        <br></br>
        <p class="w3-margin-left hover-underline-animation"><a class="w3-text-black" style="font-size: 30px;text-decoration: none;" href="{}">Fuente: Ministerio de Sanidad {}</a></p>
    <script src='https://cdnjs.cloudflare.com/ajax/libs/jquery/3.1.0/jquery.min.js'></script>
    <script src='https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-alpha.6/js/bootstrap.min.js'></script>
    </body>
    </html>
    """.format(dropdown, get_dropdown(data), html_part, WEB, date)

    return html


def get_dropdown(data):
    html_part = ""
    i = 0
    data_ordered = data[0:20].copy()
    data_ordered.sort()

    for a in range(4):
        for b in range(5):
            aux = data_ordered[i][0]
            html_part += """
                        <div class="dropdown-col">
                            <a class='dropdown-item' href='{}.html'>{}</a>
                        </div>""".format(aux, aux)
            i += 1

    return html_part


def get_top_list(rows_list, right=False):
    html_part = ""
    color = "w3-green"
    if right:
        color = "w3-light-green"

    for row in rows_list:
        html_part += """
        <div class="row">
            <div class="col-md-2" align="right">
                <a href="{}.html">
                    <img src="img/{}.jpg" style="width:60%">
                </a>
            </div>
            <div class="col-md-1 align-self-center">
                <h1 class="w3-text-black" style="font-family: Lovelo">{}º</h1>
            </div>
            <div class="col-md-3 align-self-center">
                <p class="w3-margin-left hover-underline-animation"><a class="w3-text-black" style="font-family: Lovelo;font-size: 30px;text-decoration: none;" href="{}.html">{}</a></p>
            </div>
            <div class="col-md-5 align-self-center" align="left">
                <div class="w3-light-grey w3-round w3-center" style="width:100%">
                    <div class="w3-container {} w3-round w3-text-black w3-large" style="width:{:.2f}%;height:32px">{:.2f}%</div>
                </div>
            </div>
        </div>
        <br></br> 
        """.format(row[1], row[1], rows_list.index(row) + 1, row[1], row[1], color, row[0], row[0])

    return html_part


def git_push():
    repo = Repo(PATH_OF_GIT_REPO)
    repo.git.add(all=True)
    repo.index.commit('HTML update')
    origin = repo.remote(name='origin')
    origin.push()


def main():
    update_excel()
    update_html()
    git_push()


if __name__ == "__main__":
    main()
