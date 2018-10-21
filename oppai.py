import sys
import os
import pyoppai

def print_pp(acc, pp, aim_pp, speed_pp, acc_pp):
    print(
        "\n%.17g aim\n%.17g speed\n%.17g acc\n%.17g pp\nfor %.17g%%" %
            (aim_pp, speed_pp, acc_pp, pp, acc)
    )

def print_diff(stars, aim, speed):
    print(
        "\n%.17g stars\n%.17g aim stars\n%.17g speed stars" %
        (stars, aim, speed)
    )

def chk(ctx):
    err = pyoppai.err(ctx)

    if err:
        print(err)
        sys.exit(1)

def main(mod,filepath):

    # if you need to multithread, create one ctx and buffer for each thread
    ctx = pyoppai.new_ctx()

    # parse beatmap ------------------------------------------------------------
    b = pyoppai.new_beatmap(ctx)

    BUFSIZE = 2000000 # should be big enough to hold the .osu file
    buf = pyoppai.new_buffer(BUFSIZE)

    pyoppai.parse(
        filepath,
        b,
        buf,
        BUFSIZE,

        # don't disable caching and use python script's folder for caching
        False,
        os.path.dirname(os.path.realpath(__file__))
    );

    chk(ctx)
    if not mod:
        # diff calc ----------------------------------------------------------------
        dctx = pyoppai.new_d_calc_ctx(ctx)

        stars, aim, speed, _, _, _, _ = pyoppai.d_calc(dctx, b)
        chk(ctx)


        # pp calc ------------------------------------------------------------------
        acc, pp, aim_pp, speed_pp, acc_pp = \
                pyoppai.pp_calc(ctx, aim, speed, b)

        chk(ctx)


        # pp calc (with acc %) -----------------------------------------------------
        acc, pp95, aim_pp, speed_pp, acc_pp = \
            pyoppai.pp_calc_acc(ctx, aim, speed, b, 95.0)

        chk(ctx)

        result=[stars,pp,pp95]
        return (result)
    
    else:
        dctx = pyoppai.new_d_calc_ctx(ctx)
        # mods are a bitmask, same as what the osu! api uses
        mods = pyoppai.dt
        pyoppai.apply_mods(b, mods)

        # mods are map-changing, recompute diff
        stars, aim, speed, _, _, _, _ = pyoppai.d_calc(dctx, b)
        chk(ctx)


        acc, pp, aim_pp, speed_pp, acc_pp = \
                pyoppai.pp_calc(ctx, aim, speed, b, mods)

        chk(ctx)


            # pp calc (with acc %) -----------------------------------------------------
        acc, pp95, aim_pp, speed_pp, acc_pp = \
            pyoppai.pp_calc_acc(ctx, aim, speed, b, 95.0)

        chk(ctx)

        result=[stars,pp,pp95]
        return (result)

    


