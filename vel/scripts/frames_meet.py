"""Extract frames from a meeting mp4: one every STEP seconds + scene-change frames (screen-share switches),
save full-res JPGs and contact sheets (4x3 grid, timestamps burned in). Usage: frames_meet.py <mp4> <tag> [step]"""
import sys, os, cv2, numpy as np
mp4, tag = sys.argv[1], sys.argv[2]; STEP = float(sys.argv[3]) if len(sys.argv) > 3 else 20.0
SP = os.path.dirname(os.path.abspath(__file__)); D = os.path.join(SP, "frames_" + tag); os.makedirs(D, exist_ok=True)
cap = cv2.VideoCapture(mp4); fps = cap.get(cv2.CAP_PROP_FPS); n = int(cap.get(cv2.CAP_PROP_FRAME_COUNT)); dur = n / fps
print("fps %.1f frames %d dur %.0fs" % (fps, n, dur), flush=True)
picks = []   # (t, frame)
prev_small = None; last_pick_t = -999
t = 0.0
while t < dur:
    cap.set(cv2.CAP_PROP_POS_MSEC, t * 1000); ok, fr = cap.read()
    if not ok: break
    small = cv2.resize(cv2.cvtColor(fr, cv2.COLOR_BGR2GRAY), (160, 90)).astype(np.float32)
    change = float(np.abs(small - prev_small).mean()) if prev_small is not None else 999
    periodic = (t - last_pick_t) >= STEP - 0.01
    if periodic or change > 18:
        picks.append((t, fr)); last_pick_t = t
        cv2.imwrite(os.path.join(D, "f_%05d.jpg" % int(t)), fr, [cv2.IMWRITE_JPEG_QUALITY, 80])
    prev_small = small
    t += 5.0
print("picked %d frames" % len(picks), flush=True)
# contact sheets 4x3 of 480x270 thumbs
W, H, C, R = 480, 270, 4, 3
for i in range(0, len(picks), C * R):
    sheet = np.full((R * H, C * W, 3), 255, np.uint8)
    for j, (tt, fr) in enumerate(picks[i:i + C * R]):
        th = cv2.resize(fr, (W, H)); cv2.rectangle(th, (0, 0), (110, 22), (0, 0, 0), -1)
        cv2.putText(th, "%02d:%02d" % (tt // 60, tt % 60), (4, 17), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 2)
        r_, c_ = divmod(j, C); sheet[r_ * H:(r_ + 1) * H, c_ * W:(c_ + 1) * W] = th
    cv2.imwrite(os.path.join(D, "sheet_%02d.jpg" % (i // (C * R))), sheet, [cv2.IMWRITE_JPEG_QUALITY, 85])
print("contact sheets:", (len(picks) + C * R - 1) // (C * R), flush=True)
