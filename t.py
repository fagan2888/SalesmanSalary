from dataclasses import dataclass


@dataclass
class text:
    aa: str = None
    bb: str = None


t = text()

t.aa = 12
t.bb = 10

print(t)
