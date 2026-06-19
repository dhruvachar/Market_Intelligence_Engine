from langchain_ollama import ChatOllama

llm = ChatOllama(model="qwen2.5:3b", temperature=0.2)

topic = input("Enter a market: ")

prompt = f"""
You are a management consultant.

Research the following market:

{topic}

Provide:
1. Market Overview
2. Key Trends
3. Growth Drivers
4. Risks
5. Opportunities

Use bullet points.
"""

response = llm.invoke(prompt)

print(response.content)