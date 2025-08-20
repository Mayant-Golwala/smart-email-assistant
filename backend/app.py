from flask import Flask, request, jsonify
from gpt4all import GPT4All

app = Flask(__name__)
model = GPT4All("ggml-model.bin")

@app.route("/summarize", methods=["POST"])
def summarize():
    data = request.json
    results = []
    for email in data["emails"]:
        prompt = f"""
        Analyze the following email and return:
        1. A short summary
        2. The urgency level (High, Medium, Low)
        3. The recommended action (Reply, Forward, Schedule Meeting, Archive, etc.)

        Email:
        ---
        {email}
        ---
        """
        response = model.generate(prompt)
        # You can parse response if needed
        results.append({"analysis": response})
    return jsonify(results)

app.run(port=5000)