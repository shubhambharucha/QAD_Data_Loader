import { motion } from "framer-motion";

const variants = {
  hidden:  { opacity: 0, y: 16 },
  visible: { opacity: 1, y: 0  },
};

/**
 * FadeIn — wraps children in a viewport-triggered fade + slide-up.
 * delay: stagger offset in seconds.
 */
export function FadeIn({ children, delay = 0, className = "" }) {
  return (
    <motion.div
      className={className}
      variants={variants}
      initial="hidden"
      whileInView="visible"
      viewport={{ once: true, margin: "-40px" }}
      transition={{
        duration: 0.45,
        delay,
        ease: [0.16, 1, 0.3, 1], // expo out — premium feel
      }}
    >
      {children}
    </motion.div>
  );
}

/**
 * StaggerParent — wraps a list of FadeIn children and staggers them.
 * Use with FadeIn inside when you want coordinated reveal.
 */
export function StaggerParent({ children, className = "", staggerDelay = 0.07 }) {
  return (
    <motion.div
      className={className}
      initial="hidden"
      whileInView="visible"
      viewport={{ once: true, margin: "-40px" }}
      transition={{ staggerChildren: staggerDelay }}
    >
      {children}
    </motion.div>
  );
}

/**
 * StaggerChild — use inside StaggerParent.
 */
export function StaggerChild({ children, className = "" }) {
  return (
    <motion.div
      className={className}
      variants={variants}
      transition={{ duration: 0.4, ease: [0.16, 1, 0.3, 1] }}
    >
      {children}
    </motion.div>
  );
}
