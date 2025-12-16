import sys
import json
import base64

def main():
    payload = sys.stdin.read()
    data = json.loads(payload)
    # filename passed via args
    args = sys.argv
    if "--filename" in args:
        idx = args.index("--filename")
        filename = args[idx + 1]
    else:
        filename = "Success.txt"

    content = "Success!"
    encoded = base64.b64encode(content.encode("utf-8")).decode("ascii")
    # Return "filename|base64"
    print(f"{filename}|{encoded}")

if __name__ == "__main__":
    main()
