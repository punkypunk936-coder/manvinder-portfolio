document.querySelectorAll('a[href^="#"]').forEach((link) => {
  link.addEventListener("click", () => {
    link.blur();
  });
});

const memeRail = document.querySelector("#memeRail");
const memePrevious = document.querySelector("#memeRailPrevious");
const memeNext = document.querySelector("#memeRailNext");
const memeStatus = document.querySelector("#memeRailStatus");

if (memeRail && memePrevious && memeNext && memeStatus) {
  const cards = [...memeRail.querySelectorAll(".meme-card")];
  let statusFrame = null;

  const currentCardIndex = () => cards.reduce((closest, card, index) => {
    const distance = Math.abs(card.offsetLeft - memeRail.scrollLeft);
    return distance < closest.distance ? { index, distance } : closest;
  }, { index: 0, distance: Number.POSITIVE_INFINITY }).index;

  const updateRailState = () => {
    const atEnd = memeRail.scrollLeft + memeRail.clientWidth >= memeRail.scrollWidth - 2;
    const index = atEnd ? cards.length - 1 : currentCardIndex();
    memeStatus.textContent = `${index + 1} of ${cards.length}`;
    memePrevious.disabled = memeRail.scrollLeft <= 2;
    memeNext.disabled = atEnd;
  };

  const scrollToCard = (index) => {
    const card = cards[Math.max(0, Math.min(cards.length - 1, index))];
    if (!card) return;
    memeRail.scrollTo({
      left: card.offsetLeft,
      behavior: "smooth",
    });
  };

  memePrevious.addEventListener("click", () => scrollToCard(currentCardIndex() - 1));
  memeNext.addEventListener("click", () => scrollToCard(currentCardIndex() + 1));
  memeRail.addEventListener("scroll", () => {
    window.cancelAnimationFrame(statusFrame);
    statusFrame = window.requestAnimationFrame(updateRailState);
  }, { passive: true });
  window.addEventListener("resize", updateRailState);
  updateRailState();
}
