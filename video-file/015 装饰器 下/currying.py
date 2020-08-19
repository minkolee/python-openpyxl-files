import random


def f(x, y, z):
    return 3 * x + 5 * y - 6 * z


def g(x):
    def h(y):
        def i(z):
            return 3 * x + 5 * y - 6 * z

        return i

    return h


for j in range(0, 100):
    a = random.randint(1, 100)
    b = random.randint(1, 100)
    c = random.randint(1, 100)

    print("x={} y={} z={} f(x, y, z)={} g(x)(y)(z)={} Result: {}".format(a, b, c, f(a, b, c), g(a)(b)(c),
                                                                         f(a, b, c) == g(a)(b)(c)))
