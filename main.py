from pptx import Presentation
from pptx.shapes.placeholder import LayoutPlaceholder, PlaceholderPicture
import pandas as pd
import copy
from math import ceil
def get_data():
    # return a pandas dataframe with 80 random rows. Each row represent a person at a conference. Each person as a name, a surname, a University.
    data = {
        "name": ["John", "Jane", "Alice", "Bob", "Charlie", "David", "Eve", "Frank", "Grace", "Hannah", "Ivan", "Jack", "Katie", "Liam", "Mia", "Nathan", "Olivia", "Peter", "Quinn", "Rachel", "Steve", "Tina", "Umberto", "Violet", "Walter", "Xavier", "Yvonne", "Zach", "Albert", "Bella", "Carmen", "Dylan", "Ella", "Fiona", "George", "Helen", "Igor", "Jenny", "Karl", "Lara", "Micheal", "Nora", "Oscar", "Pamela", "Quentin", "Rita", "Sam", "Tara", "Ugo", "Valeria", "William", "Xena", "Yuri", "Zoe", "Alessandro", "Beatrice", "Carlo", "Davide", "Elena", "Fabio", "Giulia", "Hugo", "Irene", "Jorge", "Klara", "Luca", "Maria", "Nico", "Ottavia", "Paolo", "Quirino", "Riccardo", "Sara", "Tommaso", "Umberto", "Valentina", "Walter", "Xena", "Yuri", "Zoe"],
        "surname": ["Doe", "Smith", "Brown", "Johnson", "Williams", "Jones", "Garcia", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "White", "Harris", "Martin", "Thompson", "Garcia", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "White", "Harris", "Martin", "Thompson", "Garcia", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "White", "Harris", "Martin", "Thompson", "Garcia", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "White", "Harris", "Martin", "Thompson", "Garcia", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "White", "Harris", "Martin"],
        "university": ["University of Banana", "University of Apple", "University of Pear", "University of Orange", "University of Strawberry", "University of Cherry", "University of Kiwi", "University of Pineapple", "University of Lemon", "University of Blueberry", "University of Raspberry", "University of Blackberry", "University of Watermelon", "University of Melon", "University of Grape", "University of Peach", "University of Plum", "University of Apricot", "University of Mango", "University of Papaya", "University of Guava", "University of Pomegranate", "University of Coconut", "University of Kiwi", "University of Pineapple", "University of Lemon", "University of Blueberry", "University of Raspberry", "University of Blackberry", "University of Watermelon", "University of Melon", "University of Grape", "University of Peach", "University of Plum", "University of Apricot", "University of Mango", "University of Papaya", "University of Guava", "University of Pomegranate", "University of Coconut", "University of Kiwi", "University of Pineapple", "University of Lemon", "University of Blueberry", "University of Raspberry", "University of Blackberry", "University of Watermelon", "University of Melon", "University of Grape", "University of Peach", "University of Plum", "University of Apricot", "University of Mango", "University of Papaya", "University of Guava", "University of Pomegranate", "University of Coconut", "University of Kiwi", "University of Pineapple", "University of Lemon", "University of Blueberry", "University of Raspberry", "University of Blackberry", "University of Watermelon", "University of Melon", "University of Grape", "University of Peach", "University of Plum", "University of Apricot", "University of Mango", "University of Papaya", "University of Guava", "University of Pomegranate", "University of Coconut", "University of Kiwi", "University of Pineapple", "University of Lemon", "University of Blueberry", "University of Atheneum", "Unversity of Lyceum"]
    }
    return pd.DataFrame(data)


# p = "./test.pptx"
# p = "./badge 1.pptx"
p = "./badge-layout.pptx"

def duplicate_slide(pres, slide_index):
    template = prs.slide_layouts[slide_index]
    new_slide = pres.slides.add_slide(template)

    # Copy the elements from the template slide to the new slide
    for shape in template.shapes:
        if shape.has_text_frame:
            new_shape = new_slide.shapes.add_shape(
                shape.auto_shape_type, shape.left, shape.top, shape.width, shape.height
            )
            new_shape.text = shape.text_frame.text
        elif shape.has_chart:
            # For charts, you might need to copy data and properties explicitly
            # Here we're keeping it simple and skipping charts
            pass
        else:
            # For other types of shapes, you may need to handle them similarly
            pass

    return new_slide



def copy_slide_from_external_prs(prs):

    # copy from external presentation all objects into the existing presentation
    external_pres = Presentation(p)

    # specify the slide you want to copy the contents from
    ext_slide = external_pres.slides[0]

    # Define the layout you want to use from your generated pptx
    SLD_LAYOUT = 6
    slide_layout = prs.slide_layouts[SLD_LAYOUT]

    # create now slide, to copy contents to 
    curr_slide = prs.slides.add_slide(slide_layout)

    # now copy contents from external slide, but do not copy slide properties
    # e.g. slide layouts, etc., because these would produce errors, as diplicate
    # entries might be generated

    for shp in ext_slide.shapes:
        el = shp.element
        newel = copy.deepcopy(el)
        curr_slide.shapes._spTree.insert_element_before(newel, 'p:extLst')

    return prs

if __name__ == '__main__':
    data = get_data()
    n = len(data)
    
    print(f"Dataframe with {n} rows")


    N_BADGE_PER_SLIDE = 2
    N_SLIDE_TO_DUPLICATE = ceil((n-N_BADGE_PER_SLIDE) / N_BADGE_PER_SLIDE)

    # The powerpoint should have only 1 slide with the template.
    prs = Presentation(p)

    # Duplicate the slide as many times as needed
    # e.g. 2 badge for slide, 10 people, 5 slides, but 1 already exists, so you will need to duplicate the template 4 times
    for i in range(N_SLIDE_TO_DUPLICATE):
        prs = copy_slide_from_external_prs(prs)

    # The placeholder values. Careful to the spaces, they are important. It's better an exact match due to "name/suNAME" or "nome/cogNOME"
    PLACEHOLDER_NAME = "Nome "
    PLACEHOLDER_SURNAME = "COGNOME"
    PLACEHOLDER_UNIVERSITY = "University of Banana"

    K_NAME = "name"
    K_SURNAME = "surname"
    K_UNIVERSITY = "university"

    idx_person=0
    c=0 # In this template, there are 3 placeholders, so I need to increment the index of the person only when I have changed all the placeholders. c is a counter for this purpose

    # For each slide
    for slide in prs.slides:
        # Check all the elements (a Shape can be nested in a group of shapes, so I need to check all the shapes in the slide)
        for shape in slide.shapes:
            # If there is not text, it's a group of shapes (old template)
            if not shape.has_text_frame:
                print("No text frame")

                if type(shape) == PlaceholderPicture:
                    continue

                #### OLD TEMPLATE ###
                # Per ogni shape del gruppo (che è un badge)
                if shape.shapes:
                    print("\tBut shapes")
                    # Per ogni elemento del badge
                    for s in shape.shapes:
                        # Se c'è un elemento di testo
                        if s.has_text_frame:
                            for paragraph in s.text_frame.paragraphs:
                                # Controllo se ho finito le persone (capita con un numero dispari di persone se i badge sono 2 per slide)
                                if idx_person >= n:
                                    break
                                # Prendi i dati della persona
                                person = data.iloc[idx_person]

                                # Un paragraph può avere più run
                                for run in paragraph.runs:
                                    print("\t\t Before|",run.text,'|')
                                    # Se è un placeholder, cambia il testo con i dati della persona
                                    if  PLACEHOLDER_NAME == run.text:
                                        run.text = person["name"]
                                    elif PLACEHOLDER_SURNAME == run.text:
                                        run.text = person["surname"]
                                    elif PLACEHOLDER_UNIVERSITY in run.text:
                                        run.text = person["university"]
                                    print("\t\t After|",run.text,'|')
                # Finito un badge, passa alla persona successiva
                idx_person += 1
                ### END OLD TEMPLATE ###
            else:
                print("Has text frame")
                s = shape # Alias for the shape

                # Check all the paragraphs in the text_frame
                for paragraph in s.text_frame.paragraphs:
                    # Controllo se ho finito le persone (capita con un numero dispari di persone se i badge sono 2 per slide)
                    if idx_person >= n:
                        break

                    # Prendi i dati della persona corrente
                    person = data.iloc[idx_person]

                    # Un paragraph può avere più run. Cosa sia un run non l'ho ancora capito. Forse una frase del paragrafo? Comunque un sottoinsieme del paragrafo
                    for run in paragraph.runs:
                        print("\t\t Before|",run.text,'|')

                        # Se è un placeholder, cambia il testo con i dati della persona
                        if  PLACEHOLDER_NAME == run.text:
                            run.text = person[K_NAME]
                            c+=1
                        elif PLACEHOLDER_SURNAME == run.text:
                            run.text = person[K_SURNAME]
                            c+=1
                        elif PLACEHOLDER_UNIVERSITY in run.text:
                            run.text = person[K_UNIVERSITY]
                            c+=1

                        print("\t\t After|",run.text,'|')
                # It checked all the paragraphs in this shape
                if c ==3:
                    idx_person += 1
                    c=0

    # for slide in prs.slides:
    #     for shape in slide.shapes:
    #         # All'inizio si entra nella shape gruppo
    #         if not shape.has_text_frame:
    #             print("No text frame")

    #             if type(shape) == PlaceholderPicture:
    #                 continue
    #             # Per ogni shape del gruppo (che è un badge)
    #             if shape.shapes:
    #                 print("\tBut shapes")
    #                 # Per ogni elemento del badge
    #                 for s in shape.shapes:
    #                     # Se c'è un elemento di testo
    #                     if s.has_text_frame:
    #                         for paragraph in s.text_frame.paragraphs:
    #                             # Controllo se ho finito le persone (capita con un numero dispari di persone se i badge sono 2 per slide)
    #                             if idx_row >= n:
    #                                 break
    #                             # Prendi i dati della persona
    #                             person = data.iloc[idx_row]

    #                             # Un paragraph può avere più run
    #                             for run in paragraph.runs:
    #                                 print("\t\t Before|",run.text,'|')
    #                                 # Se è un placeholder, cambia il testo con i dati della persona
    #                                 if  K_NAME == run.text:
    #                                     run.text = person["name"]
    #                                 elif K_SURNAME == run.text:
    #                                     run.text = person["surname"]
    #                                 elif K_UNIVERSITY in run.text:
    #                                     run.text = person["university"]
    #                                 print("\t\t After|",run.text,'|')
    #             # Finito un badge, passa alla persona successiva
    #             idx_row += 1



    prs.save('testMODIFICATO.pptx')
