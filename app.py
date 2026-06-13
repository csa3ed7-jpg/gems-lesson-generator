from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import google.generativeai as genai
import os
import json
import tempfile

app = Flask(__name__)
CORS(app)

genai.configure(api_key=os.environ.get("GEMINI_API_KEY"))
model = genai.GenerativeModel("gemini-2.5-flash")

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "AKIS_Eng.pptx")

SWIM = {
    1:'Pool rules & safe entry/exit', 2:'Water confidence & submersion',
    3:'Supine back float', 4:'Prone front float', 5:'Flutter kick on back',
    6:'Flutter kick on front', 7:'Arm action — catch, pull, recovery',
    8:'Arms + kick combined', 9:'Rotational breathing', 10:'Full stroke 5m',
    11:'Body position supine', 12:'Flutter kick on back',
    13:'Arm action backstroke', 14:'Arms + kick combined BS',
    15:'Full backstroke 5m', 16:'Front crawl 10m', 17:'Backstroke 10m',
    18:'End of Unit Assessment'
}

PACING = {
    "Grade 1": {
        1:  {"unit":"Back to School","std":"Introductory Week","lo":"follow PE class rules and demonstrate 2 locomotor movements safely"},
        2:  {"unit":"Fitness & Wellbeing","std":"1.2.1","lo":"perform a locomotor circuit (walk/run/hop/jump) with control"},
        3:  {"unit":"Fitness & Wellbeing","std":"1.2.1","lo":"perform locomotor skills with changes in direction and speed"},
        4:  {"unit":"Fitness & Wellbeing","std":"1.2.2","lo":"balance on different body parts and hold for 3 seconds"},
        5:  {"unit":"Fitness & Wellbeing","std":"2.2.6","lo":"explain in simple words why being active helps the heart and muscles"},
        6:  {"unit":"Fitness & Wellbeing","std":"2.2.9","lo":"identify healthy snack and drink choices that give energy for PE"},
        7:  {"unit":"Fitness & Wellbeing","std":"1.2.2","lo":"combine locomotor and non-locomotor skills in a sequence"},
        8:  {"unit":"Fitness & Wellbeing","std":"2.2.11","lo":"set a simple goal and track progress in a physical activity"},
        9:  {"unit":"Manipulative Skills - Feet","std":"1.2.11","lo":"dribble a large ball with inside of foot keeping it close"},
        10: {"unit":"Manipulative Skills - Feet","std":"1.2.11","lo":"change direction around cones while dribbling slowly"},
        11: {"unit":"Manipulative Skills - Feet","std":"1.2.11","lo":"receive a rolling ball and stop it with sole, then pass"},
        12: {"unit":"Manipulative Skills - Hands","std":"1.2.6","lo":"dribble a ball with two hands while standing with controlled bounces"},
        13: {"unit":"Manipulative Skills - Hands","std":"1.2.8","lo":"catch a gentle pass with ready hands then protect ball"},
        14: {"unit":"Manipulative Skills - Hands","std":"1.2.8","lo":"chest pass with a light ball to a partner from short distance"},
        15: {"unit":"Manipulative Skills - Hands","std":"1.2.8","lo":"combine throw and catch in a cooperative activity"},
        16: {"unit":"Cooperative Activities","std":"2.2.8","lo":"play a simple 3v3 game focusing on sharing and space"},
        17: {"unit":"Cooperative Activities","std":"2.2.8","lo":"demonstrate sharing, taking turns and encouraging a partner"},
        18: {"unit":"Kids Athletics","std":"1.2.3","lo":"run with high knees and arm swing in short sprints"},
        19: {"unit":"Kids Athletics","std":"1.2.3","lo":"jump for distance with two-foot take-off and soft landing"},
        20: {"unit":"Kids Athletics","std":"1.2.3","lo":"demonstrate sprint start and run 20m with control"},
        21: {"unit":"Gymnastics","std":"1.2.5","lo":"create a 3-move sequence: shape, balance, jump land"},
        22: {"unit":"Gymnastics","std":"1.2.5","lo":"balance on different body parts and hold for 3 seconds"},
        23: {"unit":"Gymnastics","std":"1.2.5","lo":"create a movement sequence using 3 different locomotor skills"},
        24: {"unit":"Gymnastics","std":"1.2.5","lo":"perform a sequence with smooth transitions"},
        25: {"unit":"Gymnastics","std":"1.2.5","lo":"demonstrate balance with control and body awareness"},
        26: {"unit":"Gymnastics","std":"1.2.3","lo":"jump for distance with two-foot take-off and soft landing"},
        27: {"unit":"Gymnastics","std":"1.2.5","lo":"create and perform a gymnastics sequence with a partner"},
        28: {"unit":"Gymnastics","std":"1.2.5","lo":"perform a gymnastics sequence independently with control"},
        29: {"unit":"Gymnastics","std":"1.2.5","lo":"evaluate and improve a gymnastics sequence"},
        30: {"unit":"Review","std":"1.2.1","lo":"demonstrate locomotor skills with changes in speed and direction"},
        31: {"unit":"Review","std":"1.2.2","lo":"combine locomotor and non-locomotor skills in a sequence"},
        32: {"unit":"Review","std":"1.2.4","lo":"demonstrate rolling and catching a ball with two hands"},
        33: {"unit":"Review","std":"1.2.3","lo":"demonstrate sprint start and run 20m with control"},
        34: {"unit":"Review","std":"2.2.6","lo":"explain why daily physical activity is important for health"},
        35: {"unit":"End of Year","std":"1.2.1","lo":"demonstrate favourite PE skills learned this year"},
        36: {"unit":"End of Year","std":"1.2.1","lo":"demonstrate favourite PE skills and celebrate growth"},
        37: {"unit":"End of Year","std":"2.2.6","lo":"reflect on physical activity progress and set a goal for next year"},
        38: {"unit":"End of Year","std":"2.2.6","lo":"participate in end-of-year PE celebrations and skill showcase"},
        39: {"unit":"End of Year","std":"2.2.6","lo":"complete end-of-year PE assessments and reflections"},
    },
    "Grade 2": {
        1:  {"unit":"Back to School","std":"Introductory Week","lo":"demonstrate PE protocols and personal space awareness"},
        2:  {"unit":"Fitness & Wellbeing","std":"1.2.1","lo":"combine 2 locomotor skills in a sequence with control"},
        3:  {"unit":"Fitness & Wellbeing","std":"1.2.1","lo":"travel using different pathways: straight, curved, zigzag"},
        4:  {"unit":"Fitness & Wellbeing","std":"1.2.5","lo":"balance on different body parts and hold for 3 seconds"},
        5:  {"unit":"Fitness & Wellbeing","std":"2.2.6","lo":"link activities to fitness: heart, muscles, flexibility"},
        6:  {"unit":"Fitness & Wellbeing","std":"2.2.7","lo":"explain cool-down and perform 3 static stretches correctly"},
        7:  {"unit":"Fitness & Wellbeing","std":"2.2.9","lo":"describe effort using simple scale and choose appropriate pace"},
        8:  {"unit":"Fitness & Wellbeing","std":"2.2.11","lo":"choose healthy fuel and hydration timing for PE day"},
        9:  {"unit":"Manipulative Skills - Feet","std":"1.2.11","lo":"dribble with inside foot in straight line keeping ball within 1 step"},
        10: {"unit":"Manipulative Skills - Feet","std":"1.2.11","lo":"dribble using inside foot through zig-zag pathway"},
        11: {"unit":"Manipulative Skills - Feet","std":"1.2.10","lo":"shoot with instep at goal target from short distance"},
        12: {"unit":"Manipulative Skills - Hands","std":"1.2.6","lo":"dribble with one hand at waist height while stationary"},
        13: {"unit":"Manipulative Skills - Hands","std":"1.2.7","lo":"perform forward roll with chin tucked and controlled finish"},
        14: {"unit":"Manipulative Skills - Hands","std":"1.2.8","lo":"chest pass and bounce pass choice to partner"},
        15: {"unit":"Manipulative Skills - Hands","std":"1.2.9","lo":"standing long jump with consistent take-off and landing"},
        16: {"unit":"Kids Athletics","std":"1.2.3","lo":"sprint with start position and accelerate for 20m"},
        17: {"unit":"Kids Athletics","std":"1.2.3","lo":"standing long jump with consistent take-off and landing"},
        18: {"unit":"Kids Athletics","std":"1.2.3","lo":"demonstrate running technique: arms, posture and stride"},
        19: {"unit":"Kids Athletics","std":"1.2.3","lo":"perform standing long jump with two-foot takeoff and landing"},
        20: {"unit":"Kids Athletics","std":"1.2.3","lo":"perform sprint start and run 30m with control"},
        21: {"unit":"Gymnastics","std":"1.2.5","lo":"create a 4-move sequence including roll and balance with smooth transitions"},
        22: {"unit":"Gymnastics","std":"1.2.5","lo":"perform forward roll with chin tucked and controlled finish"},
        23: {"unit":"Gymnastics","std":"1.2.3","lo":"perform weight transfer on multiple body parts"},
        24: {"unit":"Gymnastics","std":"1.2.5","lo":"create and perform a gymnastics sequence with smooth transitions"},
        25: {"unit":"Gymnastics","std":"1.2.5","lo":"demonstrate balance with control and body awareness"},
        26: {"unit":"Gymnastics","std":"1.2.3","lo":"combine jump, land, balance and weight transfer in sequence"},
        27: {"unit":"Gymnastics","std":"1.2.5","lo":"perform gymnastics sequence with a partner"},
        28: {"unit":"Gymnastics","std":"1.2.5","lo":"perform gymnastics sequence independently with control"},
        29: {"unit":"Gymnastics","std":"1.2.5","lo":"evaluate and improve gymnastics sequence"},
        30: {"unit":"Review","std":"1.2.1","lo":"demonstrate locomotor skills with changes in speed and direction"},
        31: {"unit":"Review","std":"1.2.5","lo":"balance on different body parts and hold for 3 seconds"},
        32: {"unit":"Review","std":"1.2.8","lo":"demonstrate rolling and catching a ball with two hands"},
        33: {"unit":"Review","std":"1.2.3","lo":"sprint with start position and accelerate for 20m"},
        34: {"unit":"Review","std":"2.2.6","lo":"explain why daily physical activity is important for health"},
        35: {"unit":"End of Year","std":"1.2.1","lo":"demonstrate favourite PE skills learned this year"},
        36: {"unit":"End of Year","std":"1.2.1","lo":"demonstrate favourite PE skills and celebrate growth"},
        37: {"unit":"End of Year","std":"2.2.6","lo":"reflect on physical activity progress and set a goal for next year"},
        38: {"unit":"End of Year","std":"2.2.6","lo":"participate in end-of-year PE celebrations and skill showcase"},
        39: {"unit":"End of Year","std":"2.2.6","lo":"complete end-of-year PE assessments and reflections"},
    },
    "Grade 3": {
        1:  {"unit":"Back to School","std":"Introductory Week","lo":"apply class routines and demonstrate locomotor skills in warm-up"},
        2:  {"unit":"Fitness & Wellbeing","std":"1.5.1","lo":"combine varied locomotor skills in a variety of practice tasks"},
        3:  {"unit":"Fitness & Wellbeing","std":"2.5.1","lo":"use effort scale to pace a continuous activity safely"},
        4:  {"unit":"Fitness & Wellbeing","std":"1.5.3","lo":"perform weight transfer feet to hands with straight-arm support"},
        5:  {"unit":"Fitness & Wellbeing","std":"1.5.4","lo":"combine jumping, landing, rolling, balancing and transfer of weight"},
        6:  {"unit":"Fitness & Wellbeing","std":"1.5.5","lo":"combine locomotor, non-locomotor, and manipulative movements"},
        7:  {"unit":"Fitness & Wellbeing","std":"1.5.1","lo":"demonstrate locomotor skills in warm-up with control"},
        8:  {"unit":"Fitness & Wellbeing","std":"2.5.1","lo":"explain energy and hydration choices and link to performance"},
        9:  {"unit":"Manipulative Skills - Feet","std":"1.5.20","lo":"dribble with inside and outside of foot at walking pace"},
        10: {"unit":"Manipulative Skills - Feet","std":"1.5.20","lo":"use change of pace while dribbling to beat a passive defender"},
        11: {"unit":"Manipulative Skills - Feet","std":"1.5.18","lo":"shoot with instep for power and accuracy, choose target corner"},
        12: {"unit":"Manipulative Skills - Hands","std":"1.5.19","lo":"dribble with both hands and change direction (crossover intro)"},
        13: {"unit":"Manipulative Skills - Hands","std":"1.5.12","lo":"pass on the move: chest pass to moving partner"},
        14: {"unit":"Manipulative Skills - Hands","std":"1.5.17","lo":"demonstrate lay-up steps and finish without defence"},
        15: {"unit":"Manipulative Skills - Hands","std":"1.5.21","lo":"combine lay-up and set shot in non-dynamic practice tasks"},
        16: {"unit":"Kids Athletics","std":"1.5.1","lo":"sprint mechanics: drive phase and upright running"},
        17: {"unit":"Kids Athletics","std":"1.5.4","lo":"jumping: approach and take-off into sand/soft landing"},
        18: {"unit":"Kids Athletics","std":"1.5.10","lo":"demonstrate overhand throwing technique in shot put and turbo javelin"},
        19: {"unit":"Kids Athletics","std":"1.5.1","lo":"baton exchange technique in a relay with sprint running"},
        20: {"unit":"Kids Athletics","std":"1.5.1","lo":"record and compare personal best performances across sprint, jump, and throw"},
        21: {"unit":"Gymnastics","std":"1.5.2","lo":"perform weight transfer feet to hands with straight-arm support"},
        22: {"unit":"Gymnastics","std":"1.5.3","lo":"perform weight transfer feet to hands (donkey kick/bunny hop)"},
        23: {"unit":"Gymnastics","std":"1.5.4","lo":"combine jump, land, balance and weight transfer in non-dynamic environment"},
        24: {"unit":"Gymnastics","std":"1.5.4","lo":"perform final routine: roll, jump/land, balance, weight transfer"},
        25: {"unit":"Gymnastics","std":"1.5.2","lo":"demonstrate weight transfer with straight-arm support and safety"},
        26: {"unit":"Gymnastics","std":"1.5.4","lo":"combine gymnastics elements in a smooth sequence"},
        27: {"unit":"Gymnastics","std":"1.5.4","lo":"perform gymnastics sequence with partner feedback"},
        28: {"unit":"Gymnastics","std":"1.5.4","lo":"perform final gymnastics routine independently"},
        29: {"unit":"Gymnastics","std":"1.5.4","lo":"evaluate gymnastics routine against criteria"},
        30: {"unit":"Review","std":"1.5.1","lo":"combine varied locomotor skills with control"},
        31: {"unit":"Review","std":"1.5.19","lo":"dribble with both hands and change direction"},
        32: {"unit":"Review","std":"1.5.20","lo":"dribble with feet using inside and outside surface"},
        33: {"unit":"Review","std":"1.5.1","lo":"sprint mechanics and relay baton exchange"},
        34: {"unit":"Review","std":"2.5.1","lo":"explain energy and hydration choices for physical performance"},
        35: {"unit":"End of Year","std":"1.5.1","lo":"demonstrate favourite PE skills learned this year"},
        36: {"unit":"End of Year","std":"1.5.1","lo":"demonstrate favourite PE skills and celebrate growth"},
        37: {"unit":"End of Year","std":"2.5.1","lo":"reflect on physical activity progress and set a goal for next year"},
        38: {"unit":"End of Year","std":"2.5.1","lo":"participate in end-of-year PE celebrations and skill showcase"},
        39: {"unit":"End of Year","std":"2.5.1","lo":"complete end-of-year PE assessments and reflections"},
    },
    "Grade 4": {
        1:  {"unit":"Back to School","std":"Introductory Week","lo":"apply PE protocols and demonstrate spatial awareness in group activities"},
        2:  {"unit":"Health, Wellbeing & Fitness","std":"2.5.7","lo":"define the five health-related components of fitness"},
        3:  {"unit":"Health, Wellbeing & Fitness","std":"2.5.9","lo":"define the structure and purpose of an effective warm-up and cool-down"},
        4:  {"unit":"Health, Wellbeing & Fitness","std":"2.5.11","lo":"identify fitness goals set using personal fitness data applying knowledge of the FITT Principle"},
        5:  {"unit":"Health, Wellbeing & Fitness","std":"2.5.12","lo":"complete health-related fitness assessments and identify personal strengths"},
        6:  {"unit":"Health, Wellbeing & Fitness","std":"2.5.14","lo":"complete fitness assessments and identify areas for improvement"},
        7:  {"unit":"Health, Wellbeing & Fitness","std":"2.5.17","lo":"describe technology tools that support physical activity goals"},
        8:  {"unit":"Health, Wellbeing & Fitness","std":"4.5.3","lo":"describe how movement positively affects personal health"},
        9:  {"unit":"Football","std":"1.5.18","lo":"demonstrate dribbling with feet using instep, inside, and outside of the foot"},
        10: {"unit":"Football","std":"1.5.20","lo":"pass with the inside of the foot accurately over short and medium distances"},
        11: {"unit":"Football","std":"1.5.21","lo":"shoot toward goal using the instep with power and accuracy"},
        12: {"unit":"Football","std":"2.5.1","lo":"demonstrate defensive strategies in small-sided invasion games"},
        13: {"unit":"Football","std":"2.5.2","lo":"combine manipulative skills and locomotor movements in a small-sided football game"},
        14: {"unit":"Football","std":"2.5.3","lo":"demonstrate football skills and game strategies through varied small-sided games"},
        15: {"unit":"Football","std":"2.5.4","lo":"consolidate football skills through varied small-sided games with fair play"},
        16: {"unit":"Basketball","std":"1.5.12","lo":"demonstrate dribbling with hands in stationary and dynamic tasks"},
        17: {"unit":"Basketball","std":"1.5.19","lo":"demonstrate chest pass and bounce pass with correct technique"},
        18: {"unit":"Basketball","std":"1.5.21","lo":"combine lay-up and set shot using correct shooting technique"},
        19: {"unit":"Basketball","std":"2.5.1","lo":"apply man-to-man defensive positioning and help defence"},
        20: {"unit":"Basketball","std":"2.5.2","lo":"demonstrate offensive and defensive skills in small-sided basketball"},
        21: {"unit":"Basketball","std":"2.5.3","lo":"combine offensive and defensive skills in small-sided basketball games"},
        22: {"unit":"Basketball","std":"2.5.4","lo":"demonstrate offensive and defensive skills in small-sided basketball games"},
        23: {"unit":"Muscular Strength & Endurance","std":"2.5.7","lo":"define the distinction between muscular strength and muscular endurance"},
        24: {"unit":"Muscular Strength & Endurance","std":"2.5.9","lo":"define training methods that develop muscular strength and endurance"},
        25: {"unit":"Athletics","std":"1.5.1","lo":"demonstrate and apply sprinting mechanics in timed sprint practice tasks"},
        26: {"unit":"Athletics","std":"1.5.4","lo":"perform long jump and standing broad jump using consistent technique"},
        27: {"unit":"Athletics","std":"1.5.6","lo":"demonstrate overhand throwing technique in shot put and turbo javelin"},
        28: {"unit":"Athletics","std":"1.5.7","lo":"demonstrate baton exchange technique in a relay competition"},
        29: {"unit":"Athletics","std":"1.5.10","lo":"record and compare personal best performances across sprint, jump, and throw"},
        30: {"unit":"Athletics","std":"2.5.1","lo":"record and compare personal best performances and reflect on improvement"},
        31: {"unit":"Volleyball","std":"1.5.13","lo":"design and perform a volleyball-specific warm-up and cool-down"},
        32: {"unit":"Volleyball","std":"2.5.1","lo":"demonstrate forearm pass (dig) using correct platform position"},
        33: {"unit":"Volleyball","std":"1.5.17","lo":"demonstrate overhead pass (set) using finger contact and triangular hand shape"},
        34: {"unit":"Volleyball","std":"1.5.15","lo":"demonstrate overhand serve in a stationary environment"},
        35: {"unit":"Volleyball","std":"2.5.4","lo":"combine forearm pass, overhead set, and serve in a modified volleyball game"},
        36: {"unit":"Volleyball","std":"2.5.11","lo":"consolidate volleyball skills through competitive game play"},
        37: {"unit":"Volleyball","std":"2.5.13","lo":"demonstrate volleyball skills consolidated through competitive game play"},
        38: {"unit":"End of Year","std":"2.5.7","lo":"reflect on fitness progress and set goals for continued improvement"},
        39: {"unit":"End of Year","std":"2.5.13","lo":"complete end-of-year PE assessments and celebrate growth"},
    },
    "Grade 5": {
        1:  {"unit":"Back to School","std":"Introductory Week","lo":"lead warm-up routines and demonstrate PE leadership protocols"},
        2:  {"unit":"Health, Wellbeing & Fitness","std":"2.5.9","lo":"apply FITT to design a plan for one fitness component"},
        3:  {"unit":"Health, Wellbeing & Fitness","std":"2.5.9","lo":"create and evaluate a short circuit session using FITT and safety"},
        4:  {"unit":"Health, Wellbeing & Fitness","std":"2.5.10","lo":"use heart-rate or RPE to monitor intensity and reflect on effort"},
        5:  {"unit":"Health, Wellbeing & Fitness","std":"2.5.10","lo":"apply FITT to design a plan for one fitness component"},
        6:  {"unit":"Health, Wellbeing & Fitness","std":"2.5.9","lo":"create and evaluate a short circuit session using FITT and safety"},
        7:  {"unit":"Health, Wellbeing & Fitness","std":"2.5.10","lo":"use heart-rate or RPE to monitor intensity and reflect on effort"},
        8:  {"unit":"Health, Wellbeing & Fitness","std":"4.5.3","lo":"describe how movement positively affects personal health and wellbeing"},
        9:  {"unit":"Football","std":"1.5.18","lo":"execute accurate inside-foot pass over short and medium distances and receive on the move"},
        10: {"unit":"Football","std":"1.5.20","lo":"control and dribble using multiple surfaces under light pressure"},
        11: {"unit":"Football","std":"1.5.21","lo":"shoot with instep for power and accuracy and choose target corner"},
        12: {"unit":"Football","std":"2.5.1","lo":"defend using pressing triggers and compact shape"},
        13: {"unit":"Football","std":"2.5.2","lo":"combine manipulative skills and locomotor movements in a small-sided football game"},
        14: {"unit":"Football","std":"2.5.3","lo":"demonstrate football skills and game strategies through varied small-sided games"},
        15: {"unit":"Football","std":"2.5.4","lo":"consolidate football skills and game strategies through varied small-sided games"},
        16: {"unit":"Basketball","std":"1.5.12","lo":"dribble at jogging pace with both hands and change-of-pace"},
        17: {"unit":"Basketball","std":"1.5.19","lo":"execute chest and bounce pass under pressure and make quick decisions"},
        18: {"unit":"Basketball","std":"1.5.21","lo":"shoot set shot with BEEF and apply lay-up in transition"},
        19: {"unit":"Basketball","std":"2.5.1","lo":"apply man-to-man and help defence with communication"},
        20: {"unit":"Basketball","std":"2.5.2","lo":"demonstrate offensive and defensive skills in small-sided basketball"},
        21: {"unit":"Basketball","std":"2.5.3","lo":"combine offensive and defensive skills in small-sided basketball games"},
        22: {"unit":"Basketball","std":"2.5.4","lo":"demonstrate offensive and defensive skills combined in small-sided basketball games"},
        23: {"unit":"Muscular Strength & Endurance","std":"2.5.7","lo":"define muscular strength and muscular endurance with examples"},
        24: {"unit":"Muscular Strength & Endurance","std":"2.5.9","lo":"define training methods that develop muscular strength and endurance"},
        25: {"unit":"Athletics","std":"1.5.1","lo":"demonstrate sprinting mechanics: drive phase, maximum velocity, and finish"},
        26: {"unit":"Athletics","std":"1.5.4","lo":"compare standing and running long jump techniques and results"},
        27: {"unit":"Athletics","std":"1.5.10","lo":"demonstrate overhand throwing in shot put and turbo javelin events"},
        28: {"unit":"Athletics","std":"1.5.7","lo":"apply relay baton exchange technique and team strategy in timed relay"},
        29: {"unit":"Athletics","std":"1.5.1","lo":"record and compare personal best performances across sprint, jump, and throw"},
        30: {"unit":"Athletics","std":"2.5.1","lo":"record and compare personal best performances and reflect on improvement"},
        31: {"unit":"Volleyball","std":"1.5.15","lo":"design and perform a volleyball-specific warm-up and perform underhand serve"},
        32: {"unit":"Volleyball","std":"2.5.1","lo":"demonstrate forearm pass (dig) using correct platform position"},
        33: {"unit":"Volleyball","std":"1.5.17","lo":"demonstrate overhead pass (set) using finger contact and triangular hand shape"},
        34: {"unit":"Volleyball","std":"1.5.15","lo":"demonstrate overhand serve focusing on toss consistency and contact"},
        35: {"unit":"Volleyball","std":"2.5.4","lo":"combine forearm pass, overhead set, and serve in a 3v3 or 4v4 modified volleyball game"},
        36: {"unit":"Volleyball","std":"2.5.11","lo":"identify volleyball skills consolidated through competitive game play with teamwork and fair play"},
        37: {"unit":"Volleyball","std":"2.5.13","lo":"explain volleyball skills consolidated through competitive game play"},
        38: {"unit":"End of Year","std":"2.5.7","lo":"reflect on fitness progress and set goals for continued improvement"},
        39: {"unit":"End of Year","std":"2.5.13","lo":"complete end-of-year PE assessments and celebrate personal growth"},
    }
}

def get_week_data(grade, week, lesson_type):
    if lesson_type == "Swimming":
        return {"unit": "Swimming", "std": "Swimming progression", "lo": SWIM.get(week, "develop swimming technique")}
    gd = PACING.get(grade, {})
    if week in gd:
        return gd[week]
    keys = sorted(gd.keys())
    fb = keys[0] if keys else None
    for k in keys:
        if k <= week:
            fb = k
    return gd.get(fb, {"unit": "Physical Education", "std": "PE Standards", "lo": "demonstrate physical literacy skills"})

def call_gemini(grade, week, lesson_type):
    d = get_week_data(grade, week, lesson_type)
    swim_line = f"\nSwimming topic: {SWIM.get(week, '')}" if lesson_type == "Swimming" else ""

    prompt = f"""You are a GEMS Education PE lesson designer using the TLAG framework. Return ONLY a valid JSON object with no markdown fences, no ```json, no extra text.

Schema:
{{
  "title": string,
  "unit": string,
  "standard": string,
  "lo": string (full sentence starting "By the end of this lesson, students will be able to..."),
  "doNow": {{"questions":["","",""],"teacherWill":"","studentsWill":""}},
  "toKnow": {{"tier2":["","",""],"tier3":["","",""],"concepts":["","","","",""]}},
  "iDo": {{"steps":[{{"label":"Step 1","text":""}},{{"label":"Step 2","text":""}},{{"label":"Step 3","text":""}},{{"label":"Step 4","text":""}}],"thinkAloud":"","teacherWill":"","studentsWill":""}},
  "weDo": {{"activity":"","bullets":["","",""],"cfuQ":["","",""],"teacherWill":"","studentsWill":""}},
  "youDo": {{"task":"","bullets":["","",""],"decisionQ":["",""],"teacherWill":"","studentsWill":""}},
  "affirmative": {{"questions":[{{"q":"","lookFor":""}},{{"q":"","lookFor":""}},{{"q":"","lookFor":""}}],"mostUnderstood":"","someStruggled":"","mostConfused":""}},
  "exitTicket": {{"q1":"","wordBank":["","",""],"q2":""}}
}}

Generate a TLAG PE lesson.
Grade: {grade} | Week: {week} | Type: {lesson_type}
Unit: {d['unit']} | Standard: {d['std']}
LO: students will be able to {d['lo']}{swim_line}"""

    response = model.generate_content(prompt)
    text = response.text.strip()
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
    return json.loads(text)

@app.route("/")
def index():
    return send_file("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    data = request.json
    grade = data.get("grade")
    week = int(data.get("week"))
    lesson_type = data.get("type", "PE")
    try:
        lesson = call_gemini(grade, week, lesson_type)
        return jsonify({"ok": True, "lesson": lesson})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/generate-pptx", methods=["POST"])
def generate_pptx():
    data = request.json
    grade = data.get("grade")
    week = int(data.get("week"))
    lesson_type = data.get("type", "PE")
    lesson = data.get("lesson")

    if not lesson:
        try:
            lesson = call_gemini(grade, week, lesson_type)
        except Exception as e:
            return jsonify({"ok": False, "error": str(e)}), 500

    try:
        from generate_pptx import build_pptx
        tmp = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False)
        tmp.close()
        build_pptx(lesson, grade, week, lesson_type, TEMPLATE_PATH, tmp.name)
        filename = f"Lesson_{grade.replace(' ','')}_W{week}_{lesson_type}.pptx"
        return send_file(
            tmp.name,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/manifest.json")
def manifest():
    return send_file("manifest.json", mimetype="application/manifest+json")

@app.route("/service-worker.js")
def service_worker():
    return send_file("service-worker.js", mimetype="application/javascript")

@app.route("/icons/<path:filename>")
def icons(filename):
    return send_file(f"icons/{filename}", mimetype="image/png")

if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
