import win32com.client as win32
import datetime
from datetime import datetime
import logging
import random


log_erros = {}


def datas_aniversários():
    aniversariantes = {'12-03-1999':{'Aline Berte':'email.com'},
                       '12-03-1999':{'Denise Krakhecke':'email.com'},
                        '15-01-1997':{'Pedro Santos':'email.com'},
                        '10-04-2004':{'Diego Hailer':'email.com'},
                        '21-04-1967':{'Nelson Onuki':'email.com'},
                        '02-05-1988':{'Dioggo Venson':'email.com'},
                        '06-05-1996':{'Níkola Zaia':'email.com'},
                        '17-07-1991':{'João Lucas Reis':'email.com'},
                        '13-08-1999':{'Gabriel Terra':'email.com'},
                        '30-08-1997':{'Aris Gomes':'email.com'},
                        '30-08-2002':{'Guilherme Canfild':'email.com'},
                        '07-09-1974':{'Sedinei Pieta':'email.com'},
                        '16-12-1994':{'Samya Uchoa Bordallo':'email.com'}
                    }           
    return aniversariantes        

def verifica_aniversariante(aniversariantes):
    
    nomes_aniv = []
    email_aniv = []
    for data_aniversario in aniversariantes:
        for nome in aniversariantes[data_aniversario]:
            if datetime.strptime(data_aniversario,'%d-%m-%Y').strftime('%d-%m') == datetime.now().strftime('%d-%m'):
                nomes_aniv.append(nome)
                email_aniv.append(aniversariantes[data_aniversario][nome])
    return nomes_aniv, email_aniv



def envia_email(mensagem,nome,email):
    try:

        assunto = 'Seu aniversário!!'
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = email
        mail.Subject = assunto
        mail.Body = mensagem
        mail.Send()

        print(f'E-mail enviado para {nome}')
        return 200
    
    except BaseException as erro:
        logging.exception(erro)
        log_erros[datetime.now().strftime('%Y-%m-%d %H:%M')] = str(erro)
        print(f'Falha no envio do e-mail para {nome} - ({email})')
        return 999

def gera_mensagem_aniverario(nome):


    nome = nome.split()[0]
    diretores = ['Nikola, Sedinei, Nelson, Siqueira']

    if nome in diretores:
        mensagem = random.choice([f"Olá, {nome}! Hoje é o seu aniversário!!!\r\n \r\nÀ medida que você celebra mais um ano de vida, queremos aproveitar esta oportunidade para expressar nossa sincera apreciação por sua presença em nossa empresa. Seu comprometimento, ética de trabalho e habilidades excepcionais têm sido uma fonte constante de inspiração para todos nós.\r\nQue este aniversário seja um lembrete do quanto você é valorizado e amado não apenas como um colega de trabalho, mas como uma pessoa extraordinária.\r\n\r\nParabéns e que todos os seus sonhos se tornem realidade!",
                    f"Querido {nome} ! /n Hoje é um dia especial em que celebramos não apenas mais um ano de vida, mas também o impacto que vocês têm em nossas vidas e em nossa empresa. Seu compromisso, liderança e visão têm sido sempre inspiração. Que este novo ano traga ainda mais sucesso, saúde e felicidade. Que continuemos a trilhar juntos o caminho do crescimento e da realização de nossos objetivos.\r\n\r\nFeliz aniversário e aproveite seu dia!",
                    f"Hoje.. é o seu dia, {nome}!! /n No seu aniversário, queremos expressar nossa gratidão por sua orientação, apoio e inspiração. Sua dedicação e compromisso não só moldaram nossa equipe, mas também nos motivaram a alcançar novos patamares de excelência. Que cada dia traga novas oportunidades e desafios que o ajudem a crescer ainda mais como líder e como pessoa.\r\n\r\nFeliz aniversário e aproveite seu dia!",
                    f"Hey, {nome}.. Hoje é o seu dia!! /n Hoje é o seu dia especial, e queremos aproveitar esta oportunidade para reconhecer não apenas o aniversário de um líder excepcional, mas também a pessoa incrível que você é. Sua presença em nossa equipe é verdadeiramente inspiradora, e cada dia ao seu lado é uma lição de dedicação, integridade e empenho.\r\n\r\nFeliz aniversário e aproveite seu dia!"])
    else:
        mensagem = random.choice([
            f"Feliz aniversário, {nome}!\r\n\r\nEstamos muito felizes por saber que hoje é um dia muito especial..\r\nQue este novo ano da sua vida seja repleto de sucesso, realizações e momentos inesquecíveis. Hoje é o dia de celebrar não apenas o seu aniversário, mas também tudo o que você representa para nossa empresa. Sua paixão, dedicação e espírito de equipe são verdadeiramente admiráveis ​​e inspiradores.\r\n\r\nAproveite seu dia, parabens!",
            f"Hoje é o seu dia especial!\r\n\r\nParabéns {nome}, pelo seu aniversário! Que este novo ciclo traga muita alegria, saúde e prosperidade para você. Agradecemos por fazer parte da nossa equipe e por todo o seu esforço. Que este seja apenas o começo de grandes conquistas!",
            f"Feliz aniversário, {nome}!\r\n\r\nHoje é o dia de celebrar você e tudo o que você representa para nossa empresa. Sua dedicação e comprometimento são admiráveis, e estamos muito felizes por tê-lo em nossa equipe.\r\nQue seu aniversário seja repleto de felicidade, amor e sucesso. Parabéns!",
            f"Feliz aniversário, {nome}!\r\n\r\nDesejo a você um dia cheio de sorrisos, abraços e momentos inesquecíveis. Que este novo ciclo que se inicia seja repleto de amor, paz e felicidade.\r\nAproveite cada instante!"
            ])
    
    return mensagem


# while datetime.now().hour > 6 or datetime.now().hour < 20:
try:

    aniversariantes = datas_aniversários()
    nomes,lista_email = verifica_aniversariante(aniversariantes)
    if (nomes,lista_email):
        for nome,email in zip(nomes,lista_email):
            mensagem = gera_mensagem_aniverario(nome)
            retorno = envia_email(mensagem,nome,email)
            print(f'envio da mensagem: {retorno}')


except BaseException as erro:
    logging.exception(erro)
    log_erros[datetime.now().strftime('%Y-%m-%d %H:%M')] = str(erro)
