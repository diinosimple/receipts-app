import base64, pickle, io

with open("token.pickle", "rb") as f:
    data = f.read()
b64 = base64.b64encode(data).decode()
# b64 と TOKEN_PICKLE_B64 が完全一致するか確認

