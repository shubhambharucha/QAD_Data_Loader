import { motion } from "framer-motion";
import { Code2 } from "lucide-react";

export default function AnimationShowcase() {
  const animations = [
    { name: "Morphing Shapes", delay: 0 },
    { name: "Gradient Shift", delay: 0.2 },
    { name: "Floating Text", delay: 0.4 },
  ];

  return (
    <section className="py-24 px-4 sm:px-6 lg:px-8 max-w-7xl mx-auto relative">
      <motion.div
        initial={{ opacity: 0 }}
        whileInView={{ opacity: 1 }}
        transition={{ duration: 0.8 }}
        viewport={{ once: true }}
      >
        <div className="flex items-center gap-3 mb-4">
          <Code2 size={24} className="text-blue-400" />
          <h2 className="text-5xl font-black bg-gradient-to-r from-blue-400 to-purple-400 bg-clip-text text-transparent">
            Animation Showcase
          </h2>
        </div>
        <p className="text-gray-400 text-lg mb-16 max-w-2xl">
          Explore the cutting-edge animations and interactions that define modern web design.
        </p>
      </motion.div>

      {/* 3x Animation Demo Grid */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-8 mb-16">
        {/* 1. Morphing Shapes */}
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          transition={{ delay: 0.1, duration: 0.6 }}
          viewport={{ once: true }}
          className="bg-gradient-to-br from-slate-800 to-slate-900 rounded-2xl p-8 border border-slate-700 hover:border-blue-500 transition-all"
        >
          <div className="h-40 flex items-center justify-center mb-6">
            <motion.div
              animate={{
                borderRadius: ["20%", "50%", "80%", "20%"],
              }}
              transition={{ duration: 4, repeat: Infinity }}
              className="w-24 h-24 bg-gradient-to-r from-blue-500 to-cyan-500"
            />
          </div>
          <h3 className="text-xl font-bold mb-2">Morphing Shapes</h3>
          <p className="text-gray-400 text-sm">Smooth shape transformations using CSS border-radius animations.</p>
        </motion.div>

        {/* 2. Gradient Animation */}
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          transition={{ delay: 0.2, duration: 0.6 }}
          viewport={{ once: true }}
          className="bg-gradient-to-br from-slate-800 to-slate-900 rounded-2xl p-8 border border-slate-700 hover:border-purple-500 transition-all"
        >
          <div className="h-40 flex items-center justify-center mb-6">
            <motion.div
              animate={{
                backgroundPosition: ["0% 0%", "100% 100%", "0% 0%"],
              }}
              transition={{ duration: 4, repeat: Infinity }}
              style={{
                backgroundImage: "linear-gradient(45deg, #3b82f6, #8b5cf6, #ec4899, #3b82f6)",
                backgroundSize: "200% 200%",
              }}
              className="w-24 h-24 rounded-xl"
            />
          </div>
          <h3 className="text-xl font-bold mb-2">Gradient Shift</h3>
          <p className="text-gray-400 text-sm">Dynamic gradient animations that flow across your designs.</p>
        </motion.div>

        {/* 3. Floating & Scale */}
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          transition={{ delay: 0.3, duration: 0.6 }}
          viewport={{ once: true }}
          className="bg-gradient-to-br from-slate-800 to-slate-900 rounded-2xl p-8 border border-slate-700 hover:border-pink-500 transition-all"
        >
          <div className="h-40 flex items-center justify-center mb-6 relative overflow-hidden">
            <motion.div
              animate={{
                y: [0, -20, 0],
                scale: [1, 1.1, 1],
              }}
              transition={{ duration: 3, repeat: Infinity }}
              className="w-20 h-20 bg-gradient-to-r from-pink-500 to-rose-500 rounded-full shadow-lg shadow-pink-500/50"
            />
          </div>
          <h3 className="text-xl font-bold mb-2">Floating Elements</h3>
          <p className="text-gray-400 text-sm">Combine floating and scaling for engaging micro-interactions.</p>
        </motion.div>
      </div>

      {/* Featured Animation Code Snippet */}
      <motion.div
        initial={{ opacity: 0, y: 20 }}
        whileInView={{ opacity: 1, y: 0 }}
        transition={{ delay: 0.4, duration: 0.6 }}
        viewport={{ once: true }}
        className="bg-gradient-to-br from-slate-800/50 to-slate-900/50 rounded-2xl p-8 border border-slate-700 backdrop-blur-sm"
      >
        <h3 className="text-2xl font-bold mb-4 text-white">Featured Animation</h3>
        <motion.pre
          animate={{ opacity: [0.7, 1, 0.7] }}
          transition={{ duration: 4, repeat: Infinity }}
          className="bg-slate-900 p-4 rounded-lg overflow-x-auto text-sm text-green-400 font-mono"
        >
          {`@keyframes morphing {
  0% { border-radius: 20%; }
  25% { border-radius: 50%; }
  50% { border-radius: 80%; }
  75% { border-radius: 50%; }
  100% { border-radius: 20%; }
}`}
        </motion.pre>
      </motion.div>
    </section>
  );
}
