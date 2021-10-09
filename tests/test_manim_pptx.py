from manim_pptx import *
from manim import *

class TestScene(PPTXScene):
    def construct(self):

        t = Tex("Hello World!")
        self.play(Write(t, run_time=2))
        self.endSlide()
        
        c = Circle(radius=3)
        self.play(Create(c))
        d = Dot()
        d.move_to(c.get_start())
        self.play(Write(d))
        self.endSlide(autonext=True)

        self.play(MoveAlongPath(d, c))
        self.endSlide(loop=True)

        self.play(*[FadeOut(m) for m in self.mobjects])

        t2 = Tex("Bye!")
        self.play(Write(t2, run_time=1))
        self.endSlide()